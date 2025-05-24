package life.light.write;

import life.light.type.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.time.LocalDate;
import java.util.*;

class WriteFileTest {

    Map<String, TypeAccount> accounts = new HashMap<>();
    TreeSet<String> journals = new TreeSet<>();
    WriteFile writeFile = new WriteFile();

    @BeforeEach
    void setUp() {
        accounts.put("512", new TypeAccount("512", "Banque"));
        accounts.put("10500", new TypeAccount("10500", "Fond travaux"));

        journals.add("BQ");
        journals.add("JR");
    }

    @Test
    void writeFileCSVAccountsTest() {

        String filename = "." + File.separator + "TEST_temp" + File.separator + "ListeDesCompteTEST.csv";
        writeFile.writeFileCSVAccounts(accounts, filename);
        try (BufferedReader reader = new BufferedReader(new FileReader(filename))) {
            String line = reader.readLine();
            Assertions.assertEquals("Compte;Intitulé du compte;", line);
            line = reader.readLine();
            Assertions.assertEquals("10500 ; Fond travaux ; ", line);
            line = reader.readLine();
            Assertions.assertEquals("512 ; Banque ; ", line);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Erreur lors de la lecture du fichier CSV : " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelAccountsTest() {
        String filename = "." + File.separator + "TEST_temp" + File.separator + "ListeDesCompteTEST.xlsx";
        writeFile.writeFileExcelAccounts(accounts, filename);
        try (FileInputStream fileInputStream = new FileInputStream(filename);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {
            Assertions.assertEquals(1, workbook.getNumberOfSheets());
            Assertions.assertEquals("Plan comptable", workbook.getSheetAt(0).getSheetName());
            Sheet sheet = workbook.getSheetAt(0);
            Assertions.assertEquals(2, sheet.getLastRowNum());
            Row row = sheet.getRow(0);
            Assertions.assertEquals("Compte", row.getCell(0).getStringCellValue());
            Assertions.assertEquals("Intitulé du compte", row.getCell(1).getStringCellValue());
            row = sheet.getRow(1);
            Assertions.assertEquals(10500, row.getCell(0).getNumericCellValue());
            Assertions.assertEquals("Fond travaux", row.getCell(1).getStringCellValue());
            row = sheet.getRow(2);
            Assertions.assertEquals(512, row.getCell(0).getNumericCellValue());
            Assertions.assertEquals("Banque", row.getCell(1).getStringCellValue());
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Erreur lors de la lecture du fichier Excel : " + e.getMessage());
        }
    }

    @Test
    void writeFileCSVGrandLivre() {

        Line line1 = new Line("DOC001", "2023-01-15", accounts.get("512"), journals.getFirst(), accounts.get("10500"), "CHK123", "Paiement travaux", "1000.00", "");
        Line line2 = new Line("DOC002", "2023-01-20", accounts.get("10500"), journals.getFirst(), accounts.get("512"), "CHK124", "Remboursement", "", "500.00");

        TotalAccount totalAccount1 = new TotalAccount("Total compte Banque", accounts.get("512"), "1000.00", "0.00");
        TotalAccount totalAccount2 = new TotalAccount("Total compte Fond travaux", accounts.get("10500"), "0.00", "500.00");

        TotalBuilding totalBuilding = new TotalBuilding("Total général", "1000.00", "500.00");

        Object[] grandLivres = {line1, line2, totalAccount1, totalAccount2, totalBuilding};

        // Execute the method to test
        String filename = "." + File.separator + "TEST_temp" + File.separator + "GrandLivre.csv";
        writeFile.writeFileCSVGrandLivre(grandLivres, filename);

        // Verify the file content
        try (BufferedReader reader = new BufferedReader(new FileReader(filename))) {
            String line = reader.readLine();
            Assertions.assertEquals("Pièce; Date; Compte; Journal; Contrepartie; N° chèque; Libellé; Débit; Crédit;", line);
            line = reader.readLine();
            Assertions.assertEquals("DOC001 ; 2023-01-15 ; 512 ; BQ ; 10500 ; CHK123 ; Paiement travaux ; 1000.00 ;  ; ", line);
            line = reader.readLine();
            Assertions.assertEquals("DOC002 ; 2023-01-20 ; 10500 ; BQ ; 512 ; CHK124 ; Remboursement ;  ; 500.00 ; ", line);
            line = reader.readLine();
            Assertions.assertEquals(" ;  ; 512 ;  ;  ;  ; Total compte Banque ; 1000.00 ; 0.00 ; ", line);
            line = reader.readLine();
            Assertions.assertEquals(" ;  ; 10500 ;  ;  ;  ; Total compte Fond travaux ; 0.00 ; 500.00 ; ", line);
            line = reader.readLine();
            Assertions.assertEquals(" ;  ;  ;  ;  ;  ; Total général ; 1000.00 ; 500.00 ; ", line);

        } catch (IOException e) {
            e.printStackTrace();
            Assertions.fail("Error reading the CSV file: " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelGrandLivre() {

        Line line1 = new Line("000001", "2024-01-15", accounts.get("512"), journals.getFirst(), accounts.get("10500"), "CHK123", "Paiement travaux", "1000.00", "");
        Line line2 = new Line("000002", "2024-01-20", accounts.get("10500"), journals.getFirst(), accounts.get("512"), "CHK124", "Remboursement", "", "500.00");
        Line line3 = new Line("000003", "2024-01-25", accounts.get("512"), journals.getLast(), null, "CHK125", "Autre opération", "500.00", "");

        TotalAccount totalAccount1 = new TotalAccount("Total compte 1 500.00 €", accounts.get("512"), "1500.00", "0.00");
        TotalAccount totalAccount2 = new TotalAccount("Total compte 500.00 €", accounts.get("10500"), "0.00", "500.00");

        TotalBuilding totalBuilding = new TotalBuilding("Total général", "1500.00", "500.00");

        Object[] grandLivres = {line1, line2, line3, totalAccount1, totalAccount2, totalBuilding};


        // Execute the method to test
        String filename = "TestGrandLivre.xlsx";
        String filePath = "." + File.separator + "TEST_temp" + File.separator + filename;
        String pathDirectoryInvoice = "";
        writeFile.writeFileExcelGrandLivre(grandLivres, filePath, journals, pathDirectoryInvoice);

        // Verify the file content
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Check number of sheets (Grand Livre + 2 journals)
            Assertions.assertEquals(3, workbook.getNumberOfSheets());

            // Check sheet names
            Assertions.assertEquals("Grand Livre", workbook.getSheetAt(0).getSheetName());
            Assertions.assertEquals(journals.getFirst(), workbook.getSheetAt(1).getSheetName());
            Assertions.assertEquals(journals.getLast(), workbook.getSheetAt(2).getSheetName());

            // Check Grand Livre sheet content
            Sheet grandLivreSheet = workbook.getSheetAt(0);
            // Check number of rows (header + 6 data rows)
            Assertions.assertEquals(7, grandLivreSheet.getLastRowNum() + 1);

            // Check BQ journal sheet content
            Sheet bqSheet = workbook.getSheetAt(1);
            // Check number of rows (header + 2 data rows)
            Assertions.assertEquals(3, bqSheet.getLastRowNum() + 1);

            // Check JR journal sheet content
            Sheet jrSheet = workbook.getSheetAt(2);
            // Check number of rows (header + 1 data row)
            Assertions.assertEquals(2, jrSheet.getLastRowNum() + 1);

        } catch (IOException e) {
            e.printStackTrace();
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelEtatRaprochement() {

        // Create Line objects for grand livre
        List<Line> grandLivres = new ArrayList<>();
        Line line1 = new Line("000001", "2024-01-15", accounts.get("512"), journals.getFirst(), accounts.get("10500"), "CHK123", "Paiement travaux", "1000.00", "");
        grandLivres.add(line1);
        Line line2 = new Line("000002", "2024-01-20", accounts.get("512"), journals.getFirst(), accounts.get("10500"), "CHK124", "Remboursement", "", "500.00");
        grandLivres.add(line2);
        Line line3 = new Line("000003", "2024-01-25", accounts.get("512"), journals.getLast(), accounts.get("10500"), "CHK125", "Autre opération", "500.00", "");
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
        String filename = "." + File.separator + "TEST_temp" + File.separator + "EtatRapprochement_TEST.xlsx";
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
            // Check Pointage Relevé KO sheet content
            Sheet pointageReleveKOSheet = workbook.getSheetAt(1);
            // Check number of rows (header + 1 data row for unmatched entry)
            Assertions.assertEquals(2, pointageReleveKOSheet.getLastRowNum() + 1);
            // Check Pointage GL OK sheet content
            Sheet pointageGLOKSheet = workbook.getSheetAt(2);
            // Check number of rows (header + 2 data rows for matched entries)
            Assertions.assertEquals(3, pointageGLOKSheet.getLastRowNum());
            // Check Pointage GL KO sheet content
            Sheet pointageGLKOSheet = workbook.getSheetAt(3);
            // Check number of rows (header + 1 data row for unmatched entry)
            Assertions.assertEquals(2, pointageGLKOSheet.getLastRowNum() + 1);
        } catch (IOException e) {
            e.printStackTrace();
            Assertions.fail("Error reading the Excel file: " + e.getMessage());
        }
    }
}