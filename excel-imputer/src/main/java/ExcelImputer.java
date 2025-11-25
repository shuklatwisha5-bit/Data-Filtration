import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class ExcelImputer {

    private static final String DEFAULT_IMPUTATION_VALUE = "0"; // Value to replace nulls with

    public static void imputeNullsInExcel(String inputFilePath, String outputFilePath) {
        
        System.out.println("Starting imputation for: " + inputFilePath);
        
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) { // Use XSSFWorkbook for .xlsx

            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is on the first sheet
            int replacedCount = 0;

            // Iterate over all rows
            for (Row row : sheet) {
                // Iterate over all cells in the current row
                for (Cell cell : row) {
                    
                    if (isNullOrEmpty(cell)) {
                        
                        // 1. Log the detection
                        System.out.printf("   NULL/Empty detected at Row %d, Col %d. Imputing...\n",
                                          cell.getRowIndex() + 1, cell.getColumnIndex() + 1);

                        // 2. Perform replacement (Imputation)
                        replaceCellValue(cell, DEFAULT_IMPUTATION_VALUE);
                        replacedCount++;
                    }
                }
            }

            // Write the modified workbook to a new file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
            
            System.out.println("\n✅ Imputation complete!");
            System.out.println("Total values replaced: " + replacedCount);
            System.out.println("Output file saved to: " + outputFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Checks if a cell is effectively null (missing, blank, or empty string).
     */
    private static boolean isNullOrEmpty(Cell cell) {
        // 1. Check if the Cell object itself is null (can happen if using getRow().getCell(i, Row.MissingCellPolicy))
        if (cell == null) {
            return true;
        }

        // 2. Check the cell type: BLANK means the cell exists but is empty
        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }
        
        // 3. For STRING cells, check if the content is null or just whitespace
        if (cell.getCellType() == CellType.STRING) {
            String value = cell.getStringCellValue();
            return value == null || value.trim().isEmpty();
        }
        
        // 4. Handle cells that might be a formula resolving to an empty string
        if (cell.getCellType() == CellType.FORMULA) {
             try {
                 // Attempt to evaluate the formula as a string
                 String formulaResult = cell.getStringCellValue();
                 return formulaResult == null || formulaResult.trim().isEmpty();
             } catch (IllegalStateException e) {
                 // Formula might evaluate to a non-string type (e.g., number, error)
                 // We only check for empty string results here.
             }
        }

        // Otherwise, the cell has a valid value
        return false;
    }

    /**
     * Replaces the cell content with the imputation value, setting the type to Numeric or String.
     */
    private static void replaceCellValue(Cell cell, String imputationValue) {
        try {
            // Attempt to convert the imputation value to a number
            double numericValue = Double.parseDouble(imputationValue);
            cell.setCellValue(numericValue);
        } catch (NumberFormatException e) {
            // If it's not a number, set it as a string
            cell.setCellValue(imputationValue);
        }
    }

    // Main method for execution
    public static void main(String[] args) {
        String input = "student_marks_input.xlsx";
        String output = "student_marks_imputed.xlsx";
        
        // You must ensure 'student_marks_input.xlsx' exists in your project directory
        imputeNullsInExcel(input, output);
    }
}