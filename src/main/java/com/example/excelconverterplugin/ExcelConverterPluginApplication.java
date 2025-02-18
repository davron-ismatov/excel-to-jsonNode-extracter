package com.example.excelconverterplugin;

import com.example.excelconverterplugin.utils.ExcelConverter;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelConverterPluginApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelConverterPluginApplication.class, args);
        ExcelConverter.convertExcelToUserData("src\\main\\resources\\templates\\Test.xlsx");
    }

}
