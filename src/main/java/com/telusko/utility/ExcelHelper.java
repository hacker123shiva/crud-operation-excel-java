package com.telusko.utility;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.telusko.entity.Student;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelHelper {

    private static final String[] HEADERS = {"ID", "Name", "Age", "Email"};
    private static final String SHEET_NAME = "Students";

    // Create an Excel sheet
    public static void writeExcel(String filePath, List<Student> students) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET_NAME);

        // Create Header
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
        }

        // Fill data
        int rowIdx = 1;
        for (Student student : students) {
            Row row = sheet.createRow(rowIdx++);
            row.createCell(0).setCellValue(student.getId());
            row.createCell(1).setCellValue(student.getName());
            row.createCell(2).setCellValue(student.getAge());
            row.createCell(3).setCellValue(student.getEmail());
        }

        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        }
        workbook.close();
    }

    // Read Excel sheet
    public static List<Student> readExcel(String filePath) throws IOException {
        List<Student> students = new ArrayList<>();
        try (FileInputStream fileIn = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(SHEET_NAME);

            Iterator<Row> rows = sheet.iterator();
            rows.next();  // Skip header

            while (rows.hasNext()) {
                Row currentRow = rows.next();

                Student student = new Student();
                student.setId((int) currentRow.getCell(0).getNumericCellValue());
                student.setName(currentRow.getCell(1).getStringCellValue());
                student.setAge((int) currentRow.getCell(2).getNumericCellValue());
                student.setEmail(currentRow.getCell(3).getStringCellValue());

                students.add(student);
            }
        }
        return students;
    }

    // Update Excel data
    public static void updateExcel(String filePath, int id, String newName) throws IOException {
        try (FileInputStream fileIn = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(SHEET_NAME);

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;  // Skip header
                if ((int) row.getCell(0).getNumericCellValue() == id) {
                    row.getCell(1).setCellValue(newName);
                    break;
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
            workbook.close();
        }
    }

    // Delete row in Excel
    public static void deleteExcelRow(String filePath, int id) throws IOException {
        try (FileInputStream fileIn = new FileInputStream(filePath)) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(SHEET_NAME);

            int rowIndex = -1;
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;  // Skip header
                if ((int) row.getCell(0).getNumericCellValue() == id) {
                    rowIndex = row.getRowNum();
                    break;
                }
            }

            if (rowIndex != -1) {
                sheet.removeRow(sheet.getRow(rowIndex));
            }

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
            workbook.close();
        }
    }
}
