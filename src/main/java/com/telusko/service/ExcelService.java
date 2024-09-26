package com.telusko.service;



import java.io.IOException;
import java.util.List;

import com.telusko.entity.Student;
import com.telusko.utility.ExcelHelper;

public class ExcelService {
    
    private static final String FILE_PATH = "students.xlsx";
    
    // Create new Excel file and add students
    public void createExcelFile(List<Student> students) throws IOException {
        ExcelHelper.writeExcel(FILE_PATH, students);
    }

    // Read students from Excel file
    public List<Student> readExcelFile() throws IOException {
        return ExcelHelper.readExcel(FILE_PATH);
    }

    // Update student name in Excel by ID
    public void updateStudentName(int id, String newName) throws IOException {
        ExcelHelper.updateExcel(FILE_PATH, id, newName);
    }

    // Delete student from Excel by ID
    public void deleteStudentById(int id) throws IOException {
        ExcelHelper.deleteExcelRow(FILE_PATH, id);
    }
}
