package com.example.excelconverterplugin.service.impl;

import com.example.excelconverterplugin.service.ExcelExtractorService;
import com.example.excelconverterplugin.utils.ExcelConverter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Slf4j
@Service
public class ExcelExtractorServiceImpl implements ExcelExtractorService {

    @Override
    public String extractFromExcel(MultipartFile file) {
        try {
            ExcelConverter.convertExcelToJsonNode(file.getInputStream());
            return "Success";
        } catch (IOException e) {
            log.error(e.getMessage());
            return "Failure";
        }
    }
}
