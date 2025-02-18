package com.example.excelconverterplugin.service;

import org.springframework.web.multipart.MultipartFile;

public interface ExcelExtractorService {
    String extractFromExcel(MultipartFile file);
}
