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
            XSSFSheet summarySheet = summaryWorkbook.createSheet("B2_C3_Summary");

            // Set up the header for the summary sheet
            String[] headers = {"Sheet Name", "B2", "C2", "B3", "C3"};
            Row headerRow = summarySheet.createRow(0);
            for (int j = 0; j < headers.length; j++) {
                headerRow.createCell(j).setCellValue(headers[j]);
            }


            int summaryRowNum = 1;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                // Get values from cells B2, C2, B3, C3
                Row row2 = sheet.getRow(1); // Row 2 (zero-indexed)
                Row row3 = sheet.getRow(2); // Row 3 (zero-indexed)

                if (row2 != null && row3 != null) {
                    Cell b2 = row2.getCell(1); // B2
                    Cell c2 = row2.getCell(2); // C2
                    Cell b3 = row3.getCell(1); // B3
                    Cell c3 = row3.getCell(2); // C3

                    // Create a new row in the summary sheet
                    Row summaryRow = summarySheet.createRow(summaryRowNum++);
                    summaryRow.createCell(0).setCellValue(sheetName);
                    summaryRow.createCell(1).setCellValue(b2 != null ? b2.toString() : "");
                    summaryRow.createCell(2).setCellValue(c2 != null ? c2.toString() : "");
                    summaryRow.createCell(3).setCellValue(b3 != null ? b3.toString() : "");
                    summaryRow.createCell(4).setCellValue(c3 != null ? c3.toString() : "");
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
