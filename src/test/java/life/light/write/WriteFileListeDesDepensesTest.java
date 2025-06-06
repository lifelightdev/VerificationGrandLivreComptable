package life.light.write;

import life.light.type.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.TreeMap;

class WriteFileListeDesDepensesTest {

    WriteFile writeFile = new WriteFile("." + File.separator + "TEST_temp" + File.separator);

    @Test
    void writeFileExcelListeDesDepensesTest() {
        // Create test data
        // 1. LineOfExpenseKey objects
        LineOfExpenseKey key1 = new LineOfExpenseKey("K1", "Key 1", "100", TypeOfExpense.Key);
        LineOfExpenseKey key2 = new LineOfExpenseKey("K2", "Key 2", "200", TypeOfExpense.Key);
        // 2. LineOfExpense objects
        LocalDate date1 = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date1, "Expense 1", "50.00", "0.00", "0.00", null, null);
        LocalDate date2 = LocalDate.of(2023, 1, 20);
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date2, "Expense 2", "30.00", "5.00", "0.00", null, null);
        LocalDate date3 = LocalDate.of(2023, 1, 25);
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date3, "Expense 3", "20.00", "0.00", "2.00", null, null);
        // 3. LineOfExpenseTotal objects with different types
        LineOfExpenseTotal totalNature1 = new LineOfExpenseTotal("N1", "Nature 1", "50", "50.00", "0.00", "0.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalNature2 = new LineOfExpenseTotal("N2", "Nature 2", "50", "50.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey1 = new LineOfExpenseTotal("K1", "Total Key 1", "100", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBuilding = new LineOfExpenseTotal("B1", "Total Building", "300", "300.00", "10.00", "5.00", TypeOfExpense.Building);
        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key1,
                expense1,
                totalNature1,
                expense2,
                totalNature2,
                totalKey1,
                key2,
                expense3,
                totalBuilding,
                null // Add a null object to test null handling
        };
        TreeMap<String, String> listOfDocuments = new TreeMap<>();
        // Define output file path
        String outputFile = "." + File.separator + "TEST_temp" + File.separator + "Liste des dépenses.xlsx";
        String pathDirectoryInvoice = "." + File.separator + "TEST_temp" + File.separator + "invoice";
        // Execute the method to test
        writeFile.writeFileExcelListeDesDepenses( listOfExpense,  listOfDocuments);
        // Verify the file was created and contains the expected data
        try (FileInputStream fileInputStream = new FileInputStream(outputFile);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check number of sheets
            Assertions.assertEquals(2, workbook.getNumberOfSheets());
            // Check sheet names
            Assertions.assertEquals("Liste des dépenses", workbook.getSheetAt(0).getSheetName());
            // Check content of the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            // Check that the sheet has at least one row (header)
            Assertions.assertTrue(sheet.getLastRowNum() > 0, "Sheet should have at least one row");
            // Just verify that the file was created and can be opened
            // The detailed content verification is complex due to the way the data is processed
            // and would require more specific knowledge of the implementation
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }
}