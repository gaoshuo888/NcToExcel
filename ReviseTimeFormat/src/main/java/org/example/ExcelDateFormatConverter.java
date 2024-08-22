package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelDateFormatConverter {
    public static void main(String[] args) {
        String inputFilePath = "E:/DownLoad/TotalP/input.xlsx";
        String outputFilePath = "E:/DownLoad/TotalP/output00.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            SimpleDateFormat oldFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
            SimpleDateFormat newFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

            for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) { // Skip the header
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(0); // First column
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String oldDateStr = cell.getStringCellValue();
                        try {
                            Date date = oldFormat.parse(oldDateStr);
                            String newDateStr = newFormat.format(date);
                            cell.setCellValue(newDateStr);
                        } catch (ParseException e) {
                            e.printStackTrace(); // Handle the parse exception
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                workbook.write(fos);
            }

            System.out.println("Date format conversion completed successfully.");

        } catch (IOException e) {
            e.printStackTrace(); // Handle the IO exception
        }
    }
}
