package com.word.demo.controller;

import com.word.demo.service.ReadExeclService;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;

/**
 * Description: controller类
 */
@RestController
public class ReadExeclController {

    @Resource
    private ReadExeclService service;

    @ApiOperation(value = "excel导入")
    @RequestMapping(value = "/import", method = RequestMethod.POST)
    public String upLoadFile(@RequestParam("file") MultipartFile file) throws Exception {
        int result = service.excelImport(file);
        if(0 == result){
            return ("导入失败");
        }
        return ("导入成功");
    }

    @ApiOperation(value = "excel导出")
    @RequestMapping(value = "/export", method = RequestMethod.POST)
    public String export(HttpServletResponse response) throws Exception {
        service.excelExport(response);
        return ("导入成功");
    }
}
