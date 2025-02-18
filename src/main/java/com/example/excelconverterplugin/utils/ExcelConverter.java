package com.example.excelconverterplugin.utils;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelConverter {
    public static List<JsonNode> convertExcelToJsonNode(InputStream excelFile) {
        List<JsonNode> jsonNodeList = new ArrayList<>();
        List<String> keys = new ArrayList<>();

        try {
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);

            if (sheet == null) {
                throw new RuntimeException("Sheet is null");
            }

            log.info("Processing physical rows {}", sheet.getPhysicalNumberOfRows());
            processPhysicalRows(sheet, keys, jsonNodeList);
        } catch (IOException e) {
            log.error(e.getMessage());
        }

        jsonNodeList.forEach(jsonNode -> log.info(jsonNode.toString()));
        return jsonNodeList;
    }

    private static void processPhysicalRows(Sheet sheet, List<String> keys, List<JsonNode> jsonNodeList) {
        ObjectMapper objectMapper = new ObjectMapper();
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);

            if (i == 0) {
                assignKeys(row, keys);
                continue;
            }

            Map<String, Object> map = new HashMap<>();

            processValueExtraction(row, map, keys);

            JsonNode jsonNode = objectMapper.valueToTree(map);
            jsonNodeList.add(jsonNode);
        }
    }

    private static void processValueExtraction(Row row, Map<String, Object> map, List<String> keys) {
        short lastCellNum = row.getLastCellNum();

        for (short cellNum = row.getFirstCellNum(); cellNum < lastCellNum; cellNum++) {
            Cell cell = row.getCell(cellNum);

            if (cell != null)
                putValue(cell, map, keys);
        }
    }

    private static void putValue(Cell cell, Map<String, Object> map, List<String> keys) {
        switch (cell.getCellType()) {
            case STRING -> map.put(keys.get(cell.getColumnIndex()), cell.getStringCellValue());
            case NUMERIC -> map.put(keys.get(cell.getColumnIndex()), cell.getNumericCellValue());
            case BOOLEAN -> map.put(keys.get(cell.getColumnIndex()), cell.getBooleanCellValue());
            case FORMULA -> map.put(keys.get(cell.getColumnIndex()), cell.getCellFormula());
        }
    }

    private static void assignKeys(Row row, List<String> keys) {
        log.info("Assigning keys");

        row.cellIterator().forEachRemaining(cell -> {
            if (cell.getCellType().equals(CellType.BLANK) && cell.getCellType().equals(CellType._NONE)) {
                return;
            }

            String cellValue = cell.getStringCellValue().trim();

            if (cellValue.contains(" ")) {
                String[] dividedKey = cellValue.split(" ");

                for (int i = 0; i < dividedKey.length; i++) {
                    if (i == 0) {
                        dividedKey[i] = dividedKey[i].toLowerCase();
                        continue;
                    }

                    String keyPart = dividedKey[i];
                    keyPart = keyPart.replace(keyPart.charAt(0), Character.toUpperCase(keyPart.charAt(0)));
                    dividedKey[i] = keyPart;
                }

                cellValue = String.join("", dividedKey);
            } else
                cellValue = cellValue.toLowerCase();

            keys.add(cellValue);

        });
    }
}
