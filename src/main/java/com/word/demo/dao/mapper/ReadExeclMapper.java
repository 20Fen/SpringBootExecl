package com.word.demo.dao.mapper;

import com.word.demo.model.Users;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

/**
 * Description: mapper
 */
@Mapper
public interface ReadExeclMapper {

//    查询表是否存在
    Integer existTable(String tablename);
//    创建表
    Integer createTable(@Param("tableName")String tablename);
//    查询人员的id
    List<String> getAllPkid();
//    批量保存
    int insertByBatch(List<Users> usersAddList);
//    批量更新
    int updateByBatch(List<Users> usersUpdateList);
//    批量删除
    int deleteByBatch(List<Users> usersDeleteList);
}
