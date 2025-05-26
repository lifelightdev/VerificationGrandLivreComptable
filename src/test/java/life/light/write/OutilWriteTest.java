package life.light.write;

import life.light.FileOfTest;
import life.light.type.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import static life.light.FileOfTest.tempTestDir;
import static org.junit.jupiter.api.Assertions.*;

class OutilWriteTest {

    private OutilWrite outilWrite = new OutilWrite();

    @Test
    void getTotalBuilding() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        int rowNum = 1;
        Row row = sheet.createRow(rowNum++);
        List<Integer> lineTotals = new ArrayList<>();
        lineTotals.add(1);
        Line line = new Line("000001", "01/01/2024",
                new TypeAccount("401", "Fournisseur"), "Journal",
                new TypeAccount("512", "Banque"), "VIRT",
                "Label", "0.00", "0.00");
        outilWrite.getLineGrandLivre(line, row, false, null);
        row = sheet.createRow(rowNum++);
        TotalAccount totalAccount = new TotalAccount("Total 0.00",
                new TypeAccount("401", "Fournisseur"), "0.00", "0.00");
        int lastRowNumTotal = 0;
        outilWrite.getTotalAccount(totalAccount, row, lastRowNumTotal);
        row = sheet.createRow(rowNum++);

        line = new Line("000002", "01/01/2024",
                new TypeAccount("512", "Banque"), "Journal",
                new TypeAccount("401", "Fournisseur"), "VIRT",
                "Label", "0.00", "0.00");
        outilWrite.getLineGrandLivre(line, row, false, null);
        row = sheet.createRow(rowNum++);
        totalAccount = new TotalAccount("Total 0.00",
                new TypeAccount("512", "Banque"), "0.00", "0.00");
        lastRowNumTotal = rowNum;
        lineTotals.add(rowNum);
        outilWrite.getTotalAccount(totalAccount, row, lastRowNumTotal);
        row = sheet.createRow(rowNum);
        TotalBuilding totalBuilding = new TotalBuilding("Total 0.00", "0.00", "0.00");
        outilWrite.getTotalBuilding(totalBuilding, row, lineTotals);
        assertEquals("Total 0.00", workbook.getSheetAt(0).getRow(5).getCell(8).getStringCellValue());
        assertEquals("J1+J5", workbook.getSheetAt(0).getRow(5).getCell(9).getCellFormula());
        assertEquals("K1+K5", workbook.getSheetAt(0).getRow(5).getCell(10).getCellFormula());
        assertEquals("J6-K6", workbook.getSheetAt(0).getRow(5).getCell(11).getCellFormula());
    }

    @Test
    void getTotalAccount() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        int rowNum = 1;
        Row row = sheet.createRow(rowNum++);
        Line line = new Line("000001", "01/01/2024",
                new TypeAccount("401", "Fournisseur"), "Journal",
                new TypeAccount("512", "Banque"), "VIRT",
                "Label", "0.00", "0.00");
        outilWrite.getLineGrandLivre(line, row, false, null);
        row = sheet.createRow(rowNum);
        TotalAccount totalAccount = new TotalAccount("Total 0.00",
                new TypeAccount("401", "Fournisseur"), "0.00", "0.00");
        int lastRowNumTotal = 0;
        outilWrite.getTotalAccount(totalAccount, row, lastRowNumTotal);
        assertEquals(401.0, workbook.getSheetAt(0).getRow(2).getCell(0).getNumericCellValue());
        assertEquals("Fournisseur", workbook.getSheetAt(0).getRow(2).getCell(1).getStringCellValue());
        assertEquals("Total 0.00", workbook.getSheetAt(0).getRow(2).getCell(8).getStringCellValue());
        assertEquals("SUM(J2:J2)", workbook.getSheetAt(0).getRow(2).getCell(9).getCellFormula());
        assertEquals("SUM(K2:K2)", workbook.getSheetAt(0).getRow(2).getCell(10).getCellFormula());
        assertEquals("J3-K3", workbook.getSheetAt(0).getRow(2).getCell(11).getCellFormula());
    }

    @Test
    void getLineGrandLivre() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        Row row = sheet.createRow(1);
        String document = "000001";
        String date = "01/01/2024";
        Double codeAccount = 401.0;
        String labelAccount = "Fournisseur";
        String journal = "Journal";
        Double codeAccount2 = 512.0;
        String labelAccount2 = "Banque";
        String type = "VIRT";
        String label = "label";
        Double amountDebit = 100.01;
        Double amountCredit = 0.0;
        String solde = "K2-J2";
        Line line = new Line(document, date,
                new TypeAccount(Double.toString(codeAccount), labelAccount), journal,
                new TypeAccount(Double.toString(codeAccount2), labelAccount2), type,
                label, Double.toString(amountDebit), Double.toString(amountCredit));
        outilWrite.getLineGrandLivre(line, row, false, null);

        assertEquals(codeAccount, workbook.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue());
        assertEquals(labelAccount, workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        assertEquals(Double.parseDouble(document), workbook.getSheetAt(0).getRow(1).getCell(2).getNumericCellValue());
        assertEquals(date, workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue());
        assertEquals(journal, workbook.getSheetAt(0).getRow(1).getCell(4).getStringCellValue());
        assertEquals(codeAccount2, workbook.getSheetAt(0).getRow(1).getCell(5).getNumericCellValue());
        assertEquals(labelAccount2, workbook.getSheetAt(0).getRow(1).getCell(6).getStringCellValue());
        assertEquals(type, workbook.getSheetAt(0).getRow(1).getCell(7).getStringCellValue());
        assertEquals(label, workbook.getSheetAt(0).getRow(1).getCell(8).getStringCellValue());
        assertEquals(amountDebit, workbook.getSheetAt(0).getRow(1).getCell(9).getNumericCellValue());
        assertEquals(amountCredit, workbook.getSheetAt(0).getRow(1).getCell(10).getNumericCellValue());
        assertEquals(solde, workbook.getSheetAt(0).getRow(1).getCell(11).getCellFormula());
    }

    @Test
    void getLineGrandLivreReport() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        Row row = sheet.createRow(1);
        String document = "";
        String date = "01/01/2024";
        Double codeAccount = 401.0;
        String labelAccount = "Fournisseur";
        String journal = "";
        Double codeAccount2 = 0.;
        String labelAccount2 = "";
        String type = "";
        String label = "Report de 0.00€";
        Double amountDebit = 0.0;
        Double amountCredit = 0.0;
        String solde = "K2-J2";
        Line line = new Line(document, date,
                new TypeAccount(Double.toString(codeAccount), labelAccount), journal,
                new TypeAccount(codeAccount2.toString(), labelAccount2), type,
                label, amountDebit.toString(), amountCredit.toString());
        outilWrite.getLineGrandLivre(line, row, false, null);

        assertEquals(codeAccount, workbook.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue());
        assertEquals(labelAccount, workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        assertEquals(document, workbook.getSheetAt(0).getRow(1).getCell(2).getStringCellValue());
        assertEquals(date, workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue());
        assertEquals(journal, workbook.getSheetAt(0).getRow(1).getCell(4).getStringCellValue());
        assertEquals(codeAccount2, workbook.getSheetAt(0).getRow(1).getCell(5).getNumericCellValue());
        assertEquals(labelAccount2, workbook.getSheetAt(0).getRow(1).getCell(6).getStringCellValue());
        assertEquals(type, workbook.getSheetAt(0).getRow(1).getCell(7).getStringCellValue());
        assertEquals(label, workbook.getSheetAt(0).getRow(1).getCell(8).getStringCellValue());
        assertEquals(amountDebit, workbook.getSheetAt(0).getRow(1).getCell(9).getNumericCellValue());
        assertEquals(amountCredit, workbook.getSheetAt(0).getRow(1).getCell(10).getNumericCellValue());
        assertEquals(solde, workbook.getSheetAt(0).getRow(1).getCell(11).getCellFormula());
    }

    @Test
    void getLineEtatRapprochement() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Pointage");
        Row row = sheet.createRow(1);
        String document = "000001";
        String date = "01/01/2024";
        Double codeAccount = 512.0;
        String labelAccount = "Banque";
        String journal = "Journal";
        Double codeAccount2 = 401.0;
        String labelAccount2 = "Fournisseur";
        String type = "VIRT";
        String label = "label";
        Double amountDebit = 100.01;
        Double amountCredit = 0.0;
        Line line = new Line(document, date,
                new TypeAccount(Double.toString(codeAccount), labelAccount), journal,
                new TypeAccount(Double.toString(codeAccount2), labelAccount2), type,
                label, Double.toString(amountDebit), Double.toString(amountCredit));
        BankLine bankLineFound;
        Integer year = 2024;
        Integer month = 1;
        LocalDate operationDate = LocalDate.of(2024, 1, 1);
        LocalDate valueDate = LocalDate.of(2024, 1, 1);
        bankLineFound = new BankLine(year, month, operationDate, valueDate,
                new TypeAccount(Double.toString(codeAccount), labelAccount), label, amountCredit, amountDebit);
        String message = "Message";
        outilWrite.getLineEtatRapprochement(line, row, bankLineFound, message);

        assertEquals(codeAccount, workbook.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue());
        assertEquals(labelAccount, workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        assertEquals(Double.parseDouble(document), workbook.getSheetAt(0).getRow(1).getCell(2).getNumericCellValue());
        assertEquals(date, workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue());
        assertEquals(journal, workbook.getSheetAt(0).getRow(1).getCell(4).getStringCellValue());
        assertEquals(codeAccount2, workbook.getSheetAt(0).getRow(1).getCell(5).getNumericCellValue());
        assertEquals(labelAccount2, workbook.getSheetAt(0).getRow(1).getCell(6).getStringCellValue());
        assertEquals(type, workbook.getSheetAt(0).getRow(1).getCell(7).getStringCellValue());
        assertEquals(label, workbook.getSheetAt(0).getRow(1).getCell(8).getStringCellValue());
        assertEquals(amountDebit, workbook.getSheetAt(0).getRow(1).getCell(9).getNumericCellValue());
        assertEquals(amountCredit, workbook.getSheetAt(0).getRow(1).getCell(10).getNumericCellValue());

        assertEquals(year + "-" + month, workbook.getSheetAt(0).getRow(1).getCell(12).getStringCellValue());
        assertEquals(codeAccount, workbook.getSheetAt(0).getRow(1).getCell(13).getNumericCellValue());
        assertEquals(labelAccount, workbook.getSheetAt(0).getRow(1).getCell(14).getStringCellValue());
        assertEquals(operationDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")), workbook.getSheetAt(0).getRow(1).getCell(15).getStringCellValue());
        assertEquals(valueDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")), workbook.getSheetAt(0).getRow(1).getCell(16).getStringCellValue());
        assertEquals(label, workbook.getSheetAt(0).getRow(1).getCell(17).getStringCellValue());
        assertEquals(amountCredit, workbook.getSheetAt(0).getRow(1).getCell(18).getNumericCellValue());
        assertEquals(amountDebit, workbook.getSheetAt(0).getRow(1).getCell(19).getNumericCellValue());
        assertEquals(message, workbook.getSheetAt(0).getRow(1).getCell(20).getStringCellValue());
    }

    @Test
    void getCellsEnteteGrandLivre() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        outilWrite.getCellsEnteteGrandLivre(sheet);
        assertEquals("Compte", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        assertEquals("Intitulé du compte", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        assertEquals("Pièce", workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue());
        assertEquals("Date", workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue());
        assertEquals("Journal", workbook.getSheetAt(0).getRow(0).getCell(4).getStringCellValue());
        assertEquals("Contrepartie", workbook.getSheetAt(0).getRow(0).getCell(5).getStringCellValue());
        assertEquals("Intitulé de la contrepartie", workbook.getSheetAt(0).getRow(0).getCell(6).getStringCellValue());
        assertEquals("N° chèque", workbook.getSheetAt(0).getRow(0).getCell(7).getStringCellValue());
        assertEquals("Libellé", workbook.getSheetAt(0).getRow(0).getCell(8).getStringCellValue());
        assertEquals("Débit", workbook.getSheetAt(0).getRow(0).getCell(9).getStringCellValue());
        assertEquals("Crédit", workbook.getSheetAt(0).getRow(0).getCell(10).getStringCellValue());
        assertEquals("Solde (Calculé)", workbook.getSheetAt(0).getRow(0).getCell(11).getStringCellValue());
        assertEquals("Vérification", workbook.getSheetAt(0).getRow(0).getCell(12).getStringCellValue());
        assertEquals("Commentaire", workbook.getSheetAt(0).getRow(0).getCell(13).getStringCellValue());
    }

    @Test
    void getCellsEnteteEtatRapprochement() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Grand Livre");
        outilWrite.getCellsEnteteEtatRapprochement(sheet);
        assertEquals("Compte", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        assertEquals("Intitulé du compte", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        assertEquals("Pièce", workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue());
        assertEquals("Date", workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue());
        assertEquals("Journal", workbook.getSheetAt(0).getRow(0).getCell(4).getStringCellValue());
        assertEquals("Contrepartie", workbook.getSheetAt(0).getRow(0).getCell(5).getStringCellValue());
        assertEquals("Intitulé de la contrepartie", workbook.getSheetAt(0).getRow(0).getCell(6).getStringCellValue());
        assertEquals("N° chèque", workbook.getSheetAt(0).getRow(0).getCell(7).getStringCellValue());
        assertEquals("Libellé", workbook.getSheetAt(0).getRow(0).getCell(8).getStringCellValue());
        assertEquals("Débit", workbook.getSheetAt(0).getRow(0).getCell(9).getStringCellValue());
        assertEquals("Crédit", workbook.getSheetAt(0).getRow(0).getCell(10).getStringCellValue());
        assertEquals("----", workbook.getSheetAt(0).getRow(0).getCell(11).getStringCellValue());
        assertEquals("Mois du relevé", workbook.getSheetAt(0).getRow(0).getCell(12).getStringCellValue());
        assertEquals("Compte", workbook.getSheetAt(0).getRow(0).getCell(13).getStringCellValue());
        assertEquals("Intitulé du compte", workbook.getSheetAt(0).getRow(0).getCell(14).getStringCellValue());
        assertEquals("Date de l'opération", workbook.getSheetAt(0).getRow(0).getCell(15).getStringCellValue());
        assertEquals("Date de valeur", workbook.getSheetAt(0).getRow(0).getCell(16).getStringCellValue());
        assertEquals("Libellé", workbook.getSheetAt(0).getRow(0).getCell(17).getStringCellValue());
        assertEquals("Débit", workbook.getSheetAt(0).getRow(0).getCell(18).getStringCellValue());
        assertEquals("Crédit", workbook.getSheetAt(0).getRow(0).getCell(19).getStringCellValue());
        assertEquals("Commentaire", workbook.getSheetAt(0).getRow(0).getCell(20).getStringCellValue());
    }

    @Test
    void getLineOfExpense() {

        FileOfTest fileOfTest = new FileOfTest();
        try {
            fileOfTest.copyInvoiceFiles();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        // Create workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Liste des dépenses");
        Row row = sheet.createRow(1);

        // Create LineOfExpense object
        String document = "Invoice123";
        LocalDate date = LocalDate.of(2024, 1, 15);
        String label = "Achat fournitures";
        String amount = "150.50";
        String deduction = "20.00";
        String recovery = "10.00";
        LineOfExpense line = new LineOfExpense(document, date, label, amount, deduction, recovery);

        // Path to invoice directory
        String pathDirectoryInvoice = tempTestDir + File.separator + "invoice" + File.separator;


        outilWrite.getLineOfExpense(line, row, pathDirectoryInvoice);

        row = sheet.createRow(2);
        LineOfExpense line2 = new LineOfExpense("Invoice234", date, label, amount, deduction, recovery);
        outilWrite.getLineOfExpense(line2, row, pathDirectoryInvoice);

        row = sheet.createRow(3);
        LineOfExpense line3 = new LineOfExpense("Invoice345", date, label, amount, deduction, recovery);
        outilWrite.getLineOfExpense(line3, row, pathDirectoryInvoice);

        // Verify the cells
        assertEquals(document, workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        assertEquals(date.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")),
                workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        assertEquals(label, workbook.getSheetAt(0).getRow(1).getCell(2).getStringCellValue());
        assertEquals(Double.parseDouble(amount),
                workbook.getSheetAt(0).getRow(1).getCell(3).getNumericCellValue());
        assertEquals(Double.parseDouble(deduction),
                workbook.getSheetAt(0).getRow(1).getCell(4).getNumericCellValue());
        assertEquals(Double.parseDouble(recovery),
                workbook.getSheetAt(0).getRow(1).getCell(5).getNumericCellValue());

        // We can't directly test the result of getMessageFindDocument since it's private,
        // but we can verify that cell 6 has a value
        assertFalse(workbook.getSheetAt(0).getRow(1).getCell(6).getStringCellValue().contains("Impossible"));
        assertFalse(workbook.getSheetAt(0).getRow(2).getCell(6).getStringCellValue().contains("Impossible"));
        assertFalse(workbook.getSheetAt(0).getRow(3).getCell(6).getStringCellValue().contains("Impossible"));
        assertTrue(workbook.getSheetAt(0).getRow(3).getCell(6).getStringCellValue().contains("Invoice345"));
    }
}