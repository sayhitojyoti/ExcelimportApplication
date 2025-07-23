package com.Practice.ExcelImportApplication.service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class ExcelService {

    @Autowired
    private JdbcTemplate jdbcTemplate;

    public ResponseEntity<?> processExcelFile(MultipartFile file) {
        try {
            InputStream inputStream = file.getInputStream();
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return ResponseEntity.badRequest().body("No header row found.");

            List<String> columnNames = new ArrayList<>();
            Set<String> seenNames = new HashSet<>();

            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell cell = headerRow.getCell(j);
                String rawName = cell != null ? cell.getStringCellValue().trim() : "";

               
                String name = rawName.toLowerCase().replaceAll("[^a-z00-9]", "_");
                if (name.matches("^[0-9].*")) {
                    name = "col_" + name;
                }

                name = name.replaceAll("_+", "_");
                name = name.replaceAll("^_+", "").replaceAll("_+$", ""); 
                if (name.isEmpty() || seenNames.contains(name)) {
                    name = "column_" + j;
                }

                columnNames.add(name);
                seenNames.add(name);
            }

        
            String originalFileName = file.getOriginalFilename();
            String tableName = originalFileName != null
                    ? originalFileName.replaceAll("\\.xlsx?$", "").toLowerCase().replaceAll("[^a-z0-9]", "_")
                    : "default_table";

            
            jdbcTemplate.execute("DROP TABLE IF EXISTS " + tableName);

        
            StringBuilder createTableSQL = new StringBuilder("CREATE TABLE " + tableName + " (id SERIAL PRIMARY KEY");
            for (String col : columnNames) {
                createTableSQL.append(", \"").append(col).append("\" TEXT");
            }
            createTableSQL.append(")");

            jdbcTemplate.execute(createTableSQL.toString());

      
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String insertSQL = "INSERT INTO " + tableName + " (" +
                        columnNames.stream().map(col -> "\"" + col + "\"").collect(Collectors.joining(", ")) +
                        ") VALUES (" +
                        columnNames.stream().map(col -> "?").collect(Collectors.joining(", ")) +
                        ")";

                List<Object> values = new ArrayList<>();
                for (int j = 0; j < columnNames.size(); j++) {
                    Cell cell = row.getCell(j);
                    values.add(getCellValueAsString(cell));
                }

                jdbcTemplate.update(insertSQL, values.toArray());
            }

            workbook.close();
            return ResponseEntity.ok("Table created: " + tableName);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.internalServerError().body("❌ Error: " + e.getMessage());
        }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                return "";
        }
    }
}




