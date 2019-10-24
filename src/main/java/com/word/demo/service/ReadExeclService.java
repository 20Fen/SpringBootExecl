package com.word.demo.service;

import org.springframework.web.multipart.MultipartFile;

/**
 * Description: 服务接口
 */
public interface ReadExeclService {

    Integer wordImport(MultipartFile file) throws Exception;
}
