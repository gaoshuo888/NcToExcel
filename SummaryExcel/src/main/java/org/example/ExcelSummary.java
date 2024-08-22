package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelSummary {
    public static void main(String[] args) {
        String inputFile = "E:\\DownLoad\\TotalP\\output2.xlsx";
        String outputFile = "E:\\DownLoad\\TotalP\\summary.xlsx";

        try (FileInputStream fis = new FileInputStream(inputFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             XSSFWorkbook summaryWorkbook = new XSSFWorkbook()) {

            // Create a new sheet for the summary
            XSSFSheet summarySheet = summaryWorkbook.createSheet("C3 Summary");
            Row headerRow = summarySheet.createRow(0);
            headerRow.createCell(0).setCellValue("Sheet Name");
            headerRow.createCell(1).setCellValue("C3 Value");

            int summaryRowNum = 1;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                // Get the value from cell C3 (row 2, column 2 - zero-indexed)
                Row row = sheet.getRow(2);
                if (row != null) {
                    Cell cell = row.getCell(2);
                    if (cell != null) {
                        String cellValue = cell.toString();

                        // Add the sheet name and C3 value to the summary sheet
                        Row summaryRow = summarySheet.createRow(summaryRowNum++);
                        summaryRow.createCell(0).setCellValue(sheetName);
                        summaryRow.createCell(1).setCellValue(cellValue);
                    }
                }
            }

            // Write the summary to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                summaryWorkbook.write(fos);
            }

            System.out.println("Summary created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
