package life.light.write;

import life.light.type.Line;
import life.light.type.TypeAccount;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

class WriteFileCoOwnerTest {

    Map<String, TypeAccount> accounts = new HashMap<>();
    WriteFile writeFile = new WriteFile("." + File.separator + "TEST_temp" + File.separator);

    @BeforeEach
    void setUp() {
        // Regular account
        accounts.put("512", new TypeAccount("512", "Banque"));
        // Co-owner accounts
        accounts.put("45000-1", new TypeAccount("45000-1", "Co-owner 1"));
        accounts.put("45000-2", new TypeAccount("45000-2", "Co-owner 2"));
    }

    @Test
    void writeFilesExcelCoOwnerTest() {
        // Create test data
        Line line1 = new Line("DOC001", "2023-01-15", accounts.get("45000-1"), "BQ", accounts.get("512"), "CHK123", "Payment co-owner 1", "1000.00", "");
        Line line2 = new Line("DOC002", "2023-01-20", accounts.get("45000-2"), "BQ", accounts.get("512"), "CHK124", "Payment co-owner 2", "500.00", "");
        Line line3 = new Line("DOC003", "2023-01-25", accounts.get("512"), "BQ", accounts.get("45000-1"), "CHK125", "Refund co-owner 1", "", "200.00");
        Object[] grandLivres = {line1, line2, line3};
        // Define output directory and file path
        String outputDir = "." + File.separator + "TEST_temp" + File.separator;
        String pathDirectoryInvoice = "";
        // Execute the method to test
        writeFile.writeFilesExcelCoOwner(grandLivres, outputDir, accounts, pathDirectoryInvoice);
        // Verify the files were created and contain the expected data
        // Check file for co-owner 1
        String fileName1 = outputDir + "Co-owner_1.xlsx";
        try (FileInputStream fileInputStream = new FileInputStream(fileName1);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check sheet name
            Assertions.assertEquals("1", workbook.getSheetAt(0).getSheetName());
            // Check content
            Sheet sheet = workbook.getSheetAt(0);
            // Header + 1 data row (for line1) + total row
            Assertions.assertEquals(2, sheet.getLastRowNum());
            // Check data row (line1)
            Row row = sheet.getRow(1);
            Assertions.assertEquals("45000-1", row.getCell(0).getStringCellValue());
            Assertions.assertEquals("Banque", row.getCell(6).getStringCellValue());
            Assertions.assertEquals("CHK123", row.getCell(7).getStringCellValue());
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file for co-owner 1: " + e.getMessage());
        }
        // Check file for co-owner 2
        String fileName2 = outputDir + "Co-owner_2.xlsx";
        try (FileInputStream fileInputStream = new FileInputStream(fileName2);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check sheet name
            Assertions.assertEquals("2", workbook.getSheetAt(0).getSheetName());
            // Check content
            Sheet sheet = workbook.getSheetAt(0);
            // Header + 1 data row (for line2) + total row
            Assertions.assertEquals(2, sheet.getLastRowNum());
            // Check data row (line2)
            Row row = sheet.getRow(1);
            Assertions.assertEquals("45000-2", row.getCell(0).getStringCellValue());
            Assertions.assertEquals("Banque", row.getCell(6).getStringCellValue());
            Assertions.assertEquals("CHK124", row.getCell(7).getStringCellValue());
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file for co-owner 2: " + e.getMessage());
        }
    }
}