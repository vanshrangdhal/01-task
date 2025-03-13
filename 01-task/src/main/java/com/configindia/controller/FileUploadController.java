package com.configindia.controller;


import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

/**
 * File Processing Application
 * Developed by: Vansh Kumar
 * Date: [07-03-2025]
 */
@Controller
public class FileUploadController {

    @GetMapping("/")
    public String index() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadFile(@RequestParam("file") MultipartFile file,
                             @RequestParam("startRow") int startRow,
                             Model model) {
        List<List<String>> data = new ArrayList<>();
        System.out.println("File uploaded by Vansh Kumar");

        try {
            String fileName = file.getOriginalFilename();
            if (fileName != null && fileName.endsWith(".csv")) {
                data = processCSV(file.getInputStream(), startRow);
            } else if (fileName != null && (fileName.endsWith(".xls") || fileName.endsWith(".xlsx"))) {
                data = processExcel(file.getInputStream(), startRow);
            } else {
                model.addAttribute("message", "Hey, Vansh Kumar says: Please upload a valid Excel or CSV file!");
                return "upload";
            }
        } catch (Exception e) {
            model.addAttribute("message", "Error processing file: " + e.getMessage());
            return "upload";
        }

        model.addAttribute("data", data);
        return "result";
    }

    private List<List<String>> processCSV(InputStream inputStream, int startRow) throws Exception {
        List<List<String>> data = new ArrayList<>();
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8));
        String line;
        int rowIndex = 0;
        while ((line = reader.readLine()) != null) {
            if (rowIndex >= startRow) {
                String[] values = line.split(",");
                data.add(List.of(values));
            }
            rowIndex++;
        }
        return data;
    }

    private List<List<String>> processExcel(InputStream inputStream, int startRow) throws Exception {
        List<List<String>> data = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int rowIndex = 0;
        for (Row row : sheet) {
            if (rowIndex >= startRow) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(cell.toString());
                }
                data.add(rowData);
            }
            rowIndex++;
        }
        workbook.close();
        return data;
    }
}
