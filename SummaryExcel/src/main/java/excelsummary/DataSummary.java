package excelsummary;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
/**
 * FileName: ExcelSummary.java
 * 类的详细说明
 *
 * @author GaoShuo
 * @version 1.0.0
 * @Date 2024/8/22
 */
public class DataSummary {
    private String inputFile ;
    private String outputFile ;
    private String[] cellNames = {"B2", "C2", "B3" , "D2"};

    public DataSummary(String inputFile, String outputFile) {
        this.inputFile = inputFile;
        this.outputFile = outputFile;
    }

    public DataSummary() {
    }

    public String getInputFile() {
        return inputFile;
    }

    public void setInputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    public String getOutputFile() {
        return outputFile;
    }

    public void setOutputFile(String outputFile) {
        this.outputFile = outputFile;
    }

    public String[] getCellNames() {
        return cellNames;
    }

    public void setCellNames(String[] cellNames) {
        this.cellNames = cellNames;
    }

    public void dataSummary() {
        String inputFile = this.inputFile;
        String outputFile = this.outputFile;

        // Define the cell names to be summarized
        String[] cellNames = this.cellNames;

        try (FileInputStream fis = new FileInputStream(inputFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             XSSFWorkbook summaryWorkbook = new XSSFWorkbook()) {

            // Create a new sheet for the summary
            XSSFSheet summarySheet = summaryWorkbook.createSheet("Summary");

            // Set up the header for the summary sheet
            String[] headers = {"Sheet Name"};
            String[] formattedHeaders = new String[cellNames.length];
            for (int i = 0; i < cellNames.length; i++) {
                formattedHeaders[i] = cellNames[i];
            }
            String[] allHeaders = concatenate(headers, formattedHeaders);

            Row headerRow = summarySheet.createRow(0);
            for (int j = 0; j < allHeaders.length; j++) {
                headerRow.createCell(j).setCellValue(allHeaders[j]);
            }

            // Set column widths for better readability
            for (int i = 0; i < allHeaders.length; i++) {
                summarySheet.setColumnWidth(i, 4000); // Set a default column width
            }

            int summaryRowNum = 1;

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                // Extract cell values and add to summary
                addSheetSummary(summarySheet, summaryRowNum++, sheet, sheetName, cellNames);
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

    // Method to extract values and add to the summary sheet
    private static void addSheetSummary(XSSFSheet summarySheet, int rowNum, Sheet sheet, String sheetName, String[] cellNames) {
        Row summaryRow = summarySheet.createRow(rowNum);
        summaryRow.createCell(0).setCellValue(sheetName);

        for (int i = 0; i < cellNames.length; i++) {
            String cellName = cellNames[i];
            int rowIndex = Integer.parseInt(cellName.substring(1)) - 1; // Convert to zero-indexed
            int colIndex = cellName.charAt(0) - 'A'; // Convert column letter to index

            Row row = sheet.getRow(rowIndex);
            Cell cell = (row != null) ? row.getCell(colIndex) : null;

            summaryRow.createCell(i + 1).setCellValue(cell != null ? cell.toString() : "");
        }
    }

    // Utility method to concatenate two string arrays
    private static String[] concatenate(String[] a, String[] b) {
        String[] result = new String[a.length + b.length];
        System.arraycopy(a, 0, result, 0, a.length);
        System.arraycopy(b, 0, result, a.length, b.length);
        return result;
    }
}
