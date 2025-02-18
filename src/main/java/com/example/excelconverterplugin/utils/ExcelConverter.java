package com.example.excelconverterplugin.utils;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelConverter {
    private static List<String> keys = new ArrayList<>();


    public static List<JsonNode> convertExcelToUserData(String filePath) {
        List<JsonNode> jsonNodeList = new ArrayList<>();

        try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            ObjectMapper objectMapper = new ObjectMapper();
            Sheet sheet = workbook.getSheetAt(0);

            if (sheet == null) {
                throw new RuntimeException("Sheet is null");
            }

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);

                if (i == 0) {
                    assignKeys(row);
                    continue;
                }

                short lastCellNum = row.getLastCellNum();
                Map<String, Object> map = new HashMap<>();

                for (short cellNum = row.getFirstCellNum(); cellNum < lastCellNum; cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null)
                        putValue(cell, map);
                }

                JsonNode jsonNode = objectMapper.valueToTree(map);

                jsonNodeList.add(jsonNode);
            }
        } catch (IOException e) {
            log.error(e.getMessage());
        }

        jsonNodeList.forEach(jsonNode -> log.info(jsonNode.toString()));
        return jsonNodeList;
    }

    private static void putValue(Cell cell, Map<String, Object> map) {
        switch (cell.getCellType()) {
            case STRING -> map.put(keys.get(cell.getColumnIndex()), cell.getStringCellValue());
            case NUMERIC -> map.put(keys.get(cell.getColumnIndex()), cell.getNumericCellValue());
            case BOOLEAN -> map.put(keys.get(cell.getColumnIndex()), cell.getBooleanCellValue());
            case FORMULA -> map.put(keys.get(cell.getColumnIndex()), cell.getCellFormula());
        }
    }

    private static void assignKeys(Row row) {
        log.info("Assigning keys");

        row.cellIterator().forEachRemaining(cell -> {
            if (cell.getCellType().equals(CellType.BLANK) && cell.getCellType().equals(CellType._NONE)) {
                return;
            }

            String value = cell.getStringCellValue().trim();

            if (value.contains(" ")) {
                String[] split = value.split(" ");

                for (int i = 0; i < split.length; i++) {
                    if (i == 0) {
                        split[i] = split[i].toLowerCase();
                        continue;
                    }

                    String string = split[i];
                    string = string.replace(string.charAt(0), Character.toUpperCase(string.charAt(0)));
                    split[i] = string;
                }

                value = String.join("", split);
            } else
                value = value.toLowerCase();

            keys.add(value);

        });
    }
}
