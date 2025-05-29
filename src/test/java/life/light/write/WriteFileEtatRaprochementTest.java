package life.light.write;

import life.light.type.BankLine;
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
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class WriteFileEtatRaprochementTest {

    Map<String, TypeAccount> accounts = new HashMap<>();
    WriteFile writeFile = new WriteFile();

    @BeforeEach
    void setUp() {
        accounts.put("512", new TypeAccount("512", "Banque"));
        accounts.put("10500", new TypeAccount("10500", "Fond travaux"));
        accounts.put("401", new TypeAccount("401", "Un compte"));
    }

    @Test
    void writeFileExcelEtatRaprochementWithMatches() {
        // Create Line objects for grand livre
        List<Line> grandLivres = new ArrayList<>();
        Line line1 = new Line("000001", "2024-01-15", accounts.get("512"), "BQ", accounts.get("10500"), "CHK123", "Paiement travaux", "1000.00", "");
        grandLivres.add(line1);
        Line line2 = new Line("000002", "2024-01-20", accounts.get("512"), "BQ", accounts.get("10500"), "CHK124", "Remboursement", "", "500.00");
        grandLivres.add(line2);
        Line line3 = new Line("000003", "2024-01-25", accounts.get("512"), "JR", accounts.get("10500"), "CHK125", "Autre opération", "500.00", "");
        grandLivres.add(line3);
        // Create BankLine objects
        List<BankLine> bankLines = new ArrayList<>();
        LocalDate date1 = LocalDate.of(2024, 1, 15);
        BankLine bankLine1 = new BankLine(2024, 1, date1, date1, accounts.get("512"), "Paiement travaux", 0.0, 1000.0);
        bankLines.add(bankLine1);
        LocalDate date2 = LocalDate.of(2024, 1, 20);
        BankLine bankLine2 = new BankLine(2024, 1, date2, date2, accounts.get("512"), "Remboursement", 500.0, 0.0);
        bankLines.add(bankLine2);
        LocalDate date3 = LocalDate.of(2024, 1, 25);
        BankLine bankLine3 = new BankLine(2024, 1, date3, date3, accounts.get("512"), "Opération non rapprochée", 200.0, 0.0);
        bankLines.add(bankLine3);
        // Execute the method to test
        String filename = "." + File.separator + "TEST_temp" + File.separator + "EtatRapprochement_TEST1.xlsx";
        writeFile.writeFileExcelEtatRaprochement(grandLivres, filename, bankLines);
        // Verify the file content
        try (FileInputStream fileInputStream = new FileInputStream(filename);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check number of sheets (4 sheets: Pointage Relevé OK, Pointage Relevé KO, Pointage GL OK, Pointage GL KO)
            Assertions.assertEquals(4, workbook.getNumberOfSheets());
            // Check sheet names
            Assertions.assertEquals("Pointage Relevé OK", workbook.getSheetAt(0).getSheetName());
            Assertions.assertEquals("Pointage Relevé KO", workbook.getSheetAt(1).getSheetName());
            Assertions.assertEquals("Pointage GL OK", workbook.getSheetAt(2).getSheetName());
            Assertions.assertEquals("Pointage GL KO", workbook.getSheetAt(3).getSheetName());
            // Check Pointage Relevé OK sheet content
            Sheet pointageReleveOKSheet = workbook.getSheetAt(0);
            // Check number of rows (header + 2 data rows for matched entries)
            Assertions.assertEquals(3, pointageReleveOKSheet.getLastRowNum());
            // Check content of matched entries in Pointage Relevé OK
            Row row = pointageReleveOKSheet.getRow(1);
            Assertions.assertEquals(512, row.getCell(0).getNumericCellValue()); // Pièce
            Assertions.assertEquals("Fond travaux", row.getCell(6).getStringCellValue()); // Libellé
            Assertions.assertEquals("CHK123", row.getCell(7).getStringCellValue()); // Débit
            row = pointageReleveOKSheet.getRow(2);
            Assertions.assertEquals(512, row.getCell(0).getNumericCellValue()); // Pièce
            Assertions.assertEquals("Fond travaux", row.getCell(6).getStringCellValue()); // Libellé
            Assertions.assertEquals("Remboursement", row.getCell(8).getStringCellValue()); // Crédit
            // Check Pointage Relevé KO sheet content
            Sheet pointageReleveKOSheet = workbook.getSheetAt(1);
            // Check number of rows (header + 1 data row for unmatched entry)
            Assertions.assertEquals(2, pointageReleveKOSheet.getLastRowNum() + 1);
            // Check content of unmatched entry in Pointage Relevé KO
            row = pointageReleveKOSheet.getRow(1);
            Assertions.assertEquals(512, row.getCell(0).getNumericCellValue()); // Pièce
            Assertions.assertEquals("Fond travaux", row.getCell(6).getStringCellValue()); // Libellé
            Assertions.assertEquals("CHK125", row.getCell(7).getStringCellValue()); // Débit
            // Check Pointage GL OK sheet content
            Sheet pointageGLOKSheet = workbook.getSheetAt(2);
            // Check number of rows (header + 2 data rows for matched entries)
            Assertions.assertEquals(3, pointageGLOKSheet.getLastRowNum());
            // Check content of matched entries in Pointage GL OK
            row = pointageGLOKSheet.getRow(1);
            Assertions.assertEquals("Fond travaux", row.getCell(6).getStringCellValue());
            Assertions.assertEquals("Paiement travaux", row.getCell(8).getStringCellValue());
            row = pointageGLOKSheet.getRow(2);
            Assertions.assertEquals("Fond travaux", row.getCell(6).getStringCellValue());
            Assertions.assertEquals("CHK124", row.getCell(7).getStringCellValue());
            // Check Pointage GL KO sheet content
            Sheet pointageGLKOSheet = workbook.getSheetAt(3);
            // Check number of rows (header + 1 data row for unmatched entry)
            Assertions.assertEquals(2, pointageGLKOSheet.getLastRowNum() + 1);
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelEtatRaprochementWithNoMatches() {
        // Create Line objects for grand livre with no matches to bank lines
        List<Line> grandLivres = new ArrayList<>();
        Line line1 = new Line("000001", "2024-01-15", accounts.get("512"), "BQ", accounts.get("10500"), "CHK123", "Paiement travaux", "2000.00", "");
        grandLivres.add(line1);
        Line line2 = new Line("000002", "2024-01-20", accounts.get("512"), "BQ", accounts.get("10500"), "CHK124", "Remboursement", "", "1500.00");
        grandLivres.add(line2);
        // Create BankLine objects with no matches to grand livre lines
        List<BankLine> bankLines = new ArrayList<>();
        LocalDate date1 = LocalDate.of(2024, 1, 15);
        BankLine bankLine1 = new BankLine(2024, 1, date1, date1, accounts.get("512"), "Paiement travaux", 0.0, 3000.0);
        bankLines.add(bankLine1);
        LocalDate date2 = LocalDate.of(2024, 1, 20);
        BankLine bankLine2 = new BankLine(2024, 1, date2, date2, accounts.get("512"), "Remboursement", 2500.0, 0.0);
        bankLines.add(bankLine2);
        // Execute the method to test
        String filename = "." + File.separator + "TEST_temp" + File.separator + "EtatRapprochement_TEST2.xlsx";
        writeFile.writeFileExcelEtatRaprochement(grandLivres, filename, bankLines);
        // Verify the file content
        try (FileInputStream fileInputStream = new FileInputStream(filename);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check number of sheets
            Assertions.assertEquals(4, workbook.getNumberOfSheets());
            // Check Pointage Relevé OK sheet content (should be empty except for header)
            Sheet pointageReleveOKSheet = workbook.getSheetAt(0);
            Assertions.assertEquals(1, pointageReleveOKSheet.getLastRowNum());
            // Check Pointage Relevé KO sheet content (should have all grand livre entries)
            Sheet pointageReleveKOSheet = workbook.getSheetAt(1);
            Assertions.assertEquals(3, pointageReleveKOSheet.getLastRowNum() + 1);
            // Check Pointage GL OK sheet content (should be empty except for header)
            Sheet pointageGLOKSheet = workbook.getSheetAt(2);
            Assertions.assertEquals(1, pointageGLOKSheet.getLastRowNum());
            // Check Pointage GL KO sheet content (should have all bank line entries)
            Sheet pointageGLKOSheet = workbook.getSheetAt(3);
            Assertions.assertEquals(3, pointageGLKOSheet.getLastRowNum() + 1);
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelEtatRaprochementWithDifferentAccountTypes() {
        // Create Line objects for grand livre with different account types
        List<Line> grandLivres = new ArrayList<>();
        Line line1 = new Line("000001", "2024-01-15", accounts.get("512"), "BQ", accounts.get("10500"), "CHK123", "Paiement travaux", "1000.00", "");
        grandLivres.add(line1);
        Line line2 = new Line("000002", "2024-01-20", accounts.get("401"), "BQ", accounts.get("10500"), "CHK124", "Facture fournisseur", "500.00", "");
        grandLivres.add(line2);
        // Create BankLine objects
        List<BankLine> bankLines = new ArrayList<>();
        LocalDate date1 = LocalDate.of(2024, 1, 15);
        BankLine bankLine1 = new BankLine(2024, 1, date1, date1, accounts.get("512"), "Paiement travaux", 0.0, 1000.0);
        bankLines.add(bankLine1);
        LocalDate date2 = LocalDate.of(2024, 1, 20);
        BankLine bankLine2 = new BankLine(2024, 1, date2, date2, accounts.get("401"), "Facture fournisseur", 0.0, 500.0);
        bankLines.add(bankLine2);
        // Execute the method to test
        String filename = "." + File.separator + "TEST_temp" + File.separator + "EtatRapprochement_TEST3.xlsx";
        writeFile.writeFileExcelEtatRaprochement(grandLivres, filename, bankLines);
        // Verify the file content
        try (FileInputStream fileInputStream = new FileInputStream(filename);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            // Check Pointage Relevé OK sheet content
            Sheet pointageReleveOKSheet = workbook.getSheetAt(0);
            Assertions.assertEquals(2, pointageReleveOKSheet.getLastRowNum());
            // Check Pointage GL OK sheet content
            Sheet pointageGLOKSheet = workbook.getSheetAt(2);
            Assertions.assertEquals(2, pointageGLOKSheet.getLastRowNum());
        } catch (IOException e) {
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }
}