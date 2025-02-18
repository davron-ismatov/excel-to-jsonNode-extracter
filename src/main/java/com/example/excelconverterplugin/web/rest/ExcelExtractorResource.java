package com.example.excelconverterplugin.web.rest;

import com.example.excelconverterplugin.service.ExcelExtractorService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/api/excel-extractor")
public class ExcelExtractorResource {
    private final ExcelExtractorService service;

    @PostMapping("/clients-info")
    public String clients(@RequestPart MultipartFile file) {
        log.info("Clients info came in file : {}", file.getOriginalFilename());
        return service.extractFromExcel(file);
    }
}
