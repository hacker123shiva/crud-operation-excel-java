package com.telusko;

 

import com.telusko.entity.Student;
import com.telusko.service.ExcelService;

import java.util.Arrays;
import java.util.List;


public class Main {
    
    public static void main(String[] args) {
      

        ExcelService service = new ExcelService();

        List<Student> students = Arrays.asList(
            new Student(1, "Shiva", 20, "shiva@gmail.com"),
            new Student(2, "Puchu", 21, "Puchu@gmail.com"),
            new Student(2, "Arjun", 21, "Arjun@gmail.com")
        );

        try {
            // Create and write to Excel file
            service.createExcelFile(students);

            // Read data from Excel
            List<Student> readStudents = service.readExcelFile();
            readStudents.forEach(student -> System.out.println(student.getName()));

            // Update a student's name
            service.updateStudentName(2, "Bandriya");

            // Delete a student
            service.deleteStudentById(2);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}