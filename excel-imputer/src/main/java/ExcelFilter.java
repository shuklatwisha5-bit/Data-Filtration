package com.data.filter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFilter {

    // Input and Output file names
    private static final String INPUT_FILE = "student_marks_input.xlsx";
    private static final String OUTPUT_FILE = "science_high_scorers_output.xlsx";
    // Column index for Science score (Column D, 0-indexed is 3)
    private static final int SCORE_COLUMN_INDEX = 3;
    private static final double MINIMUM_SCORE = 75.0; // New minimum score for Science

    public static void main(String[] args) {
        System.out.println("Starting data filtering for: " + INPUT_FILE);
        System.out.println("Filtering rule: Science score (Column D) >= " + MINIMUM_SCORE + ".");

        try {
            // Read the input workbook
            Workbook inputWorkbook = readWorkbook(INPUT_FILE);
            Sheet inputSheet = inputWorkbook.getSheetAt(0);

            // Create a brand new workbook for the filtered results
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Science High Scorers");

            int outputRowNum = 0;
            int filteredRowCount = 0;

            // 1. Copy Header Row (Row 0)
            Row headerRow = inputSheet.getRow(0);
            if (headerRow != null) {
                // Copy the header row to the new sheet
                copyRow(headerRow, outputSheet.createRow(outputRowNum++));
            }

            // 2. Iterate through data rows (starting from row 1)
            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row currentRow = inputSheet.getRow(i);
                if (currentRow == null) continue;

                // Get the target score cell (Column D, index 3 for Science)
                Cell targetCell = currentRow.getCell(SCORE_COLUMN_INDEX);

                // Check if the cell exists and contains a number
                if (targetCell != null && targetCell.getCellType() == CellType.NUMERIC) {
                    double score = targetCell.getNumericCellValue();

                    // Apply the filtering condition
                    if (score >= MINIMUM_SCORE) {
                        // Copy the entire row to the output sheet
                        copyRow(currentRow, outputSheet.createRow(outputRowNum++));
                        filteredRowCount++;
                    }
                }
            }

            // Write the filtered changes to the new output file
            writeWorkbook(outputWorkbook, OUTPUT_FILE);

            System.out.println("\n--- FILTERING COMPLETE ---");
            System.out.println("Successfully processed data and created: " + OUTPUT_FILE);
            System.out.println("Total rows added (excluding header): " + filteredRowCount);

        } catch (IOException e) {
            System.err.println("\n--- ERROR ---");
            System.err.println("Ensure '" + INPUT_FILE + "' is in the root folder and is CLOSED.");
            System.err.println(e.toString());
        } catch (Exception e) {
            System.err.println("\n--- UNEXPECTED ERROR ---");
            e.printStackTrace();
        }
    }

    // Helper method to read the workbook
    private static Workbook readWorkbook(String fileName) throws IOException {
        try (FileInputStream excelFile = new FileInputStream(fileName)) {
            return new XSSFWorkbook(excelFile);
        }
    }

    // Helper method to write the workbook
    private static void writeWorkbook(Workbook workbook, String fileName) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        }
        try {
            workbook.close();
        } catch (IOException e) {
            // Log workbook close error but continue
        }
    }

    // Helper method to copy cells from source row to destination row
    private static void copyRow(Row sourceRow, Row destinationRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell == null) continue;

            Cell destinationCell = destinationRow.createCell(i);

            // Copy cell type and value
            destinationCell.setCellType(sourceCell.getCellType());
            switch (sourceCell.getCellType()) {
                case STRING:
                    destinationCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    destinationCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    destinationCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                case ERROR:
                    destinationCell.setCellErrorValue(sourceCell.getErrorCellValue());
                    break;
                default:
                    // Do nothing for BLANK cells or other types
                    break;
            }

            // You might want to copy style as well, but for simplicity, we skip styling here.
        }
    }
}

