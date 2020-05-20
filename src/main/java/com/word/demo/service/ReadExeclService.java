package com.word.demo.service;

import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * Description: 服务接口
 */
public interface ReadExeclService {

    Integer excelImport(MultipartFile file) throws Exception;

    void excelExport(HttpServletResponse response);

}
