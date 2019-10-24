package com.word.demo.service.impl;

import com.alibaba.fastjson.JSON;
import com.word.demo.dao.mapper.ReadExeclMapper;
import com.word.demo.model.Users;
import com.word.demo.service.ReadExeclService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.annotation.Resource;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import lombok.extern.log4j.Log4j2;

/**
 * Description: 实现类
 */
@Log4j2
@Service
public class ReadExeclServiceImpl implements ReadExeclService {

    /**
     * 设备信息表名字
     */
    public static String EQUIPMENT_TABLE_NAME;

    @Value("${equipment.table.name}")
    private void setTableName(String tableName) {
        EQUIPMENT_TABLE_NAME = tableName;
    }

    @Resource
    private ReadExeclMapper mapper;

/**
 * @description   读取execl表格并到数据库
 * @author Y
 * @Param [ file ]
 * @return java.lang.Integer
 * @date 2019/10/17 19:00
 */
    @Override
    public Integer wordImport(MultipartFile file) throws Exception {
//        查询表是否存在
        int tableExist = mapper.existTable(EQUIPMENT_TABLE_NAME);
        if (tableExist < 1) {
            mapper.createTable(EQUIPMENT_TABLE_NAME);
        }
//        检查文件是否存在
        checkFile(file);
        List<String> cdList = mapper.getAllPkid();
//        操作文件
        Workbook workBook = getWorkBook(file);
        List<Map<String, Object>> addList = null;
        List<Map<String, Object>> updateList = null;
        List<Map<String, Object>> deleteList = null;
        if (workBook != null) {
            addList = new ArrayList<Map<String, Object>>();
            updateList = new ArrayList<Map<String, Object>>();
            deleteList = new ArrayList<Map<String, Object>>();
//            得到每页的页码名
            Sheet sheetAt = workBook.getSheetAt(0);
//            得到每页的行数
            int rows = sheetAt.getPhysicalNumberOfRows();
//            遵循模板规格。5行
            if (sheetAt.getRow(0) != null && sheetAt.getRow(0).getLastCellNum() == 5) {
                if (rows == 0) {
                    return null;
                }
                for (int i = 1; i < rows; i++) {
                    Map<String, Object> obj = new HashMap<String, Object>();
                    Row row = sheetAt.getRow(i);
                    //判断当前行是否为空或者当前行的前2个必填项是否为空
                    if (row == null
                            || row.getCell(0).getStringCellValue().equals("")
                            || row.getCell(1).getStringCellValue().equals("")
                    ) {
                        continue;
                    } else {
                        if (row.getLastCellNum() < 5) {
                            continue;
                        }
                        // 获取第i行第一个单元格的下标
                        int j = row.getFirstCellNum();

                        //人员id
                        String pkid = cellToString(row.getCell(j++));
                        obj.put("pkid", pkid);
                        //人员名称
                        String name = cellToString(row.getCell(j++));
                        if (name != null) {
                            obj.put("name", name);
                        } else {
                            obj.put("name", null);
                        }
                        //人员性别
                        String sex = cellToString(row.getCell(j++));
                        if (sex.equals("1") || sex.equals("2")) {
                            obj.put("sex", sex);
                        } else {
                            obj.put("sex", null);
                        }
                        //创建时间
                        String year = cellToString(row.getCell(j++));
                        if (year != null) {
                            obj.put("year", year);
                        } else {
                            obj.put("year", null);
                        }
                        String cell = cellToString(row.getCell(j++));
                        if (cell != null) {
                            if (cell.equals("√")) {
                                if (cdList.contains(cdList)) {
                                    updateList.add(obj);
                                } else {
                                    obj.put("id", UUID.randomUUID().toString());
                                    addList.add(obj);
                                }
                            } else if (cell.equals("×")) {
                                if (cdList.contains(cdList)) {
                                    deleteList.add(obj);
                                }
                            }
                        }
                        if (addList.size() + updateList.size() > 1000) {
                            break;
                        }
                    }
                }
            } else {
                throw new Exception("导入Excel模板不正确");

            }
        }

        //导入设备数量
        int importCount = addList.size() + updateList.size();
        List<Users> equipmentAddList = JSON.parseArray(JSON.toJSONString(addList), Users.class);
        List<Users> equipmentUpdateList = JSON.parseArray(JSON.toJSONString(updateList), Users.class);
        List<Users> equipmentDeleteList = JSON.parseArray(JSON.toJSONString(deleteList), Users.class);
        //添加新设备信息
        if (!equipmentAddList.isEmpty()) {
            int insertNum = mapper.insertByBatch(equipmentAddList);
        }
        //更新设备信息
        if (!equipmentUpdateList.isEmpty()) {
            int updateNum = mapper.updateByBatch(equipmentUpdateList);
        }
        //删除设备信息
        if (!equipmentDeleteList.isEmpty()) {
            int deleteNum = mapper.deleteByBatch(equipmentDeleteList);
        }
        return importCount;
    }

    // 判断文件
    public static void checkFile(MultipartFile file) throws Exception {

        if (null == file) {
            throw new Exception("文件不存在");
        }
//        获取文件名
        String filename = file.getOriginalFilename();
//        判断文件是否是execl文件
        if (!filename.endsWith("xls") && !filename.endsWith("xlsx")) {
            throw new Exception(filename + "不是excel文件");
        }
    }

    //      文件
    public static Workbook getWorkBook(MultipartFile file) throws Exception {

        if (null == file) {
            throw new Exception("文件不存在");
        }
//        获取文件名
        String filename = file.getOriginalFilename();

        Workbook workbook = null;
        try {
            InputStream is = file.getInputStream();
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if (filename.endsWith("xls")) {
                //2003版Excel
                workbook = new HSSFWorkbook(is);
            } else if (filename.endsWith("xlsx")) {
                //2007版Excel
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            throw new Exception(e.getMessage());
        }
        return workbook;
    }

    public static String cellToString(Cell cell) {
        String value = null;
        if (cell.getCellTypeEnum() == CellType.STRING) {
            value = cell.getStringCellValue();
        }
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            Object temp = cell.getNumericCellValue();
            value = temp.toString().substring(0, temp.toString().indexOf("."));
        }
        return value;
    }

}