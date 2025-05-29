package life.light.write;

import life.light.type.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeMap;

import static life.light.write.WriteOutil.NOM_ENTETE_COLONNE_GRAND_LIVRE;
import static life.light.write.WriteOutil.NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class WriteLineTest {

    private final WriteLine writeLine = new WriteLine();
    private Workbook workbook;
    private Sheet sheet;
    private TypeAccount account1;
    private TypeAccount account2;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test Sheet");
        account1 = new TypeAccount("512", "Banque");
        account2 = new TypeAccount("401", "Fournisseur");
    }

    @Test
    void getLineGrandLivre() {
        Row row = sheet.createRow(1);
        Line line = new Line("000001", "01/01/2024", account1, "Journal", account2, "VIRT", "Label", "100.00", "0.00");
        writeLine.getLineGrandLivre(line, row, false, null);
        assertEquals(512.0, row.getCell(0).getNumericCellValue());
        assertEquals("Banque", row.getCell(1).getStringCellValue());
        assertEquals(1.0, row.getCell(2).getNumericCellValue());
        assertEquals("01/01/2024", row.getCell(3).getStringCellValue());
        assertEquals("Journal", row.getCell(4).getStringCellValue());
        assertEquals(401.0, row.getCell(5).getNumericCellValue());
        assertEquals("Fournisseur", row.getCell(6).getStringCellValue());
        assertEquals("VIRT", row.getCell(7).getStringCellValue());
        assertEquals("Label", row.getCell(8).getStringCellValue());
        assertEquals(100.0, row.getCell(9).getNumericCellValue());
        assertEquals(0.0, row.getCell(10).getNumericCellValue());
    }

    @Test
    void getLineGrandLivreWithVerif() {
        Row row = sheet.createRow(1);
        Line line = new Line("000001", "01/01/2024", account1, "Journal", account2, "VIRT", "Label", "100.00", "0.00");
        writeLine.getLineGrandLivre(line, row, true, "path/to/invoices");
        assertEquals(512.0, row.getCell(0).getNumericCellValue());
        assertEquals("Banque", row.getCell(1).getStringCellValue());
        assertEquals(1.0, row.getCell(2).getNumericCellValue());
        assertEquals("01/01/2024", row.getCell(3).getStringCellValue());
        assertEquals("Journal", row.getCell(4).getStringCellValue());
        assertEquals(401.0, row.getCell(5).getNumericCellValue());
        assertEquals("Fournisseur", row.getCell(6).getStringCellValue());
        assertEquals("VIRT", row.getCell(7).getStringCellValue());
        assertEquals("Label", row.getCell(8).getStringCellValue());
        assertEquals(100.0, row.getCell(9).getNumericCellValue());
        assertEquals(0.0, row.getCell(10).getNumericCellValue());
        assertNotNull(row.getCell(12)); // Verification cell should exist
    }

    @Test
    void getLineEtatRapprochement() {
        Row row = sheet.createRow(1);
        Line line = new Line("000001", "01/01/2024", account1, "Journal", account2, "VIRT", "Label", "100.00", "0.00");
        LocalDate date = LocalDate.of(2024, 1, 1);
        BankLine bankLine = new BankLine(2024, 1, date, date, account1, "Label", 0.0, 100.0);
        String message = "Test message";
        writeLine.getLineEtatRapprochement(line, row, bankLine, message);
        // Check Grand Livre cells
        assertEquals(512.0, row.getCell(0).getNumericCellValue());
        assertEquals("Banque", row.getCell(1).getStringCellValue());
        assertEquals(1.0, row.getCell(2).getNumericCellValue());
        assertEquals("01/01/2024", row.getCell(3).getStringCellValue());
        assertEquals("Journal", row.getCell(4).getStringCellValue());
        assertEquals(401.0, row.getCell(5).getNumericCellValue());
        assertEquals("Fournisseur", row.getCell(6).getStringCellValue());
        assertEquals("VIRT", row.getCell(7).getStringCellValue());
        assertEquals("Label", row.getCell(8).getStringCellValue());
        assertEquals(100.0, row.getCell(9).getNumericCellValue());
        assertEquals(0.0, row.getCell(10).getNumericCellValue());
        // Check Bank cells
        assertEquals("2024-1", row.getCell(12).getStringCellValue());
        assertEquals(512.0, row.getCell(13).getNumericCellValue());
        assertEquals("Banque", row.getCell(14).getStringCellValue());
        assertEquals("01/01/2024", row.getCell(15).getStringCellValue());
        assertEquals("01/01/2024", row.getCell(16).getStringCellValue());
        assertEquals("Label", row.getCell(17).getStringCellValue());
        assertEquals(0.0, row.getCell(18).getNumericCellValue());
        assertEquals(100.0, row.getCell(19).getNumericCellValue());
        // Check message
        assertEquals(message, row.getCell(20).getStringCellValue());
    }

    @Test
    void getLineEtatRapprochementWithNullLine() {
        Row row = sheet.createRow(1);
        LocalDate date = LocalDate.of(2024, 1, 1);
        BankLine bankLine = new BankLine(2024, 1, date, date, account1, "Label", 0.0, 100.0);
        String message = "Test message";
        writeLine.getLineEtatRapprochement(null, row, bankLine, message);
        // Check Bank cells
        assertEquals("2024-1", row.getCell(12).getStringCellValue());
        assertEquals(512.0, row.getCell(13).getNumericCellValue());
        assertEquals("Banque", row.getCell(14).getStringCellValue());
        assertEquals("01/01/2024", row.getCell(15).getStringCellValue());
        assertEquals("01/01/2024", row.getCell(16).getStringCellValue());
        assertEquals("Label", row.getCell(17).getStringCellValue());
        assertEquals(0.0, row.getCell(18).getNumericCellValue());
        assertEquals(100.0, row.getCell(19).getNumericCellValue());
        // Check message
        assertEquals(message, row.getCell(20).getStringCellValue());
    }

    @Test
    void getLineEtatRapprochementWithNullBankLine() {
        Row row = sheet.createRow(1);
        Line line = new Line("000001", "01/01/2024", account1, "Journal", account2, "VIRT", "Label", "100.00", "0.00");
        String message = "Test message";
        writeLine.getLineEtatRapprochement(line, row, null, message);
        // Check Grand Livre cells
        assertEquals(512.0, row.getCell(0).getNumericCellValue());
        assertEquals("Banque", row.getCell(1).getStringCellValue());
        assertEquals(1.0, row.getCell(2).getNumericCellValue());
        assertEquals("01/01/2024", row.getCell(3).getStringCellValue());
        assertEquals("Journal", row.getCell(4).getStringCellValue());
        assertEquals(401.0, row.getCell(5).getNumericCellValue());
        assertEquals("Fournisseur", row.getCell(6).getStringCellValue());
        assertEquals("VIRT", row.getCell(7).getStringCellValue());
        assertEquals("Label", row.getCell(8).getStringCellValue());
        assertEquals(100.0, row.getCell(9).getNumericCellValue());
        assertEquals(0.0, row.getCell(10).getNumericCellValue());
        // Check message
        assertEquals(message, row.getCell(20).getStringCellValue());
    }

    @Test
    void getTotalBuilding() {
        Row row = sheet.createRow(5);
        List<Integer> lineTotals = new ArrayList<>();
        lineTotals.add(1);
        lineTotals.add(3);
        TotalBuilding totalBuilding = new TotalBuilding("Total général", "1000.00", "500.00");
        writeLine.getTotalBuilding(totalBuilding, row, lineTotals);
        assertEquals("Total général", row.getCell(8).getStringCellValue());
        assertEquals("J1+J3", row.getCell(9).getCellFormula());
        assertEquals("K1+K3", row.getCell(10).getCellFormula());
        assertEquals("J6-K6", row.getCell(11).getCellFormula());
    }

    @Test
    void getTotalAccount() {
        Row row = sheet.createRow(3);
        TotalAccount totalAccount = new TotalAccount("Total compte 1000.00 €", account1, "1000.00", "500.00");
        int lastRowNumTotal = 1;
        writeLine.getTotalAccount(totalAccount, row, lastRowNumTotal);
        assertEquals(512.0, row.getCell(0).getNumericCellValue());
        assertEquals("Banque", row.getCell(1).getStringCellValue());
        assertEquals("Total compte 1000.00 €", row.getCell(8).getStringCellValue());
        assertEquals("SUM(J3:J3)", row.getCell(9).getCellFormula());
        assertEquals("SUM(K3:K3)", row.getCell(10).getCellFormula());
        assertEquals("J4-K4", row.getCell(11).getCellFormula());
    }


    @Test
    void getLineOfExpenseTotal_Nature() {
        Row row = sheet.createRow(3);
        LineOfExpenseTotal line = new LineOfExpenseTotal("Nature", "Nature Label", "Nature Value", "1000.00", "500.00", "200.00", TypeOfExpense.Nature);
        int lastRowNumTotalNature = 1;
        writeLine.getLineOfExpenseTotal(line, row, lastRowNumTotalNature);
        assertEquals("Total de la nature : Nature", row.getCell(2).getStringCellValue());
        assertEquals("SUM(D1:D3)", row.getCell(3).getCellFormula());
        assertEquals("SUM(E1:E3)", row.getCell(4).getCellFormula());
        assertEquals("SUM(F1:F3)", row.getCell(5).getCellFormula());
    }

    @Test
    void getLineOfExpenseTotal_Key() {
        Row row = sheet.createRow(5);
        LineOfExpenseTotal line = new LineOfExpenseTotal("Key", "Key Label", "Key Value", "1000.00", "500.00", "200.00", TypeOfExpense.Key);
        List<Integer> listIdLineTotal = new ArrayList<>();
        listIdLineTotal.add(1);
        listIdLineTotal.add(3);
        writeLine.getLineOfExpenseTotal(line, row, listIdLineTotal);
        assertEquals("Total de la clé : Key", row.getCell(2).getStringCellValue());
        assertEquals("D1+D3", row.getCell(3).getCellFormula());
        assertEquals("E1+E3", row.getCell(4).getCellFormula());
        assertEquals("F1+F3", row.getCell(5).getCellFormula());
    }

    @Test
    void getLineOfExpenseTotal_Building() {
        Row row = sheet.createRow(7);
        LineOfExpenseTotal line = new LineOfExpenseTotal("Building", "Building Label", "Building Value", "1000.00", "500.00", "200.00", TypeOfExpense.Building);
        List<Integer> listIdLineTotal = new ArrayList<>();
        listIdLineTotal.add(1);
        listIdLineTotal.add(3);
        listIdLineTotal.add(5);
        writeLine.getLineOfExpenseTotal(line, row, listIdLineTotal);
        assertEquals("Total de l'immeuble : Building", row.getCell(2).getStringCellValue());
        assertEquals("D1+D3+D5", row.getCell(3).getCellFormula());
        assertEquals("E1+E3+E5", row.getCell(4).getCellFormula());
        assertEquals("F1+F3+F5", row.getCell(5).getCellFormula());
    }

    @Test
    void getLineOfExpenseKey() {
        Row row = sheet.createRow(1);
        LineOfExpenseKey line = new LineOfExpenseKey("Key", "Label", "Value", TypeOfExpense.Key);
        writeLine.getLineOfExpenseKey(line, row);
        assertEquals("Label : Key Value", row.getCell(2).getStringCellValue());
    }

    @Test
    void getLineOfExpense() {
        Row row = sheet.createRow(1);
        LocalDate date = LocalDate.of(2024, 1, 15);
        LineOfExpense line = new LineOfExpense("Invoice123", date, "Label", "100.00", "20.00", "10.00");
        writeLine.getLineOfExpense(line, row, "path/to/invoices");
        assertEquals("Invoice123", row.getCell(0).getStringCellValue());
        assertEquals("15/01/2024", row.getCell(1).getStringCellValue());
        assertEquals("Label", row.getCell(2).getStringCellValue());
        assertEquals(100.0, row.getCell(3).getNumericCellValue());
        assertEquals(20.0, row.getCell(4).getNumericCellValue());
        assertEquals(10.0, row.getCell(5).getNumericCellValue());
    }

    @Test
    void getListOfDocumentMissing() {
        Sheet sheetDocument = workbook.createSheet("Documents manquants");
        TreeMap<String, String> ligneOfDocumentMissing = new TreeMap<>();
        ligneOfDocumentMissing.put("123", "Not found 123");
        ligneOfDocumentMissing.put("234", "Not found 234");
        writeLine.getListOfDocumentMissing(ligneOfDocumentMissing, sheetDocument);
        assertEquals(123, sheetDocument.getRow(0).getCell(0).getNumericCellValue());
        assertEquals("Not found 123", sheetDocument.getRow(0).getCell(1).getStringCellValue());
        assertEquals(234, sheetDocument.getRow(1).getCell(0).getNumericCellValue());
        assertEquals("Not found 234", sheetDocument.getRow(1).getCell(1).getStringCellValue());
    }

    @Test
    void getCellsEnteteGrandLivre() {
        writeLine.getCellsEntete(sheet, NOM_ENTETE_COLONNE_GRAND_LIVRE);
        assertEquals("Compte", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Intitulé du compte", sheet.getRow(0).getCell(1).getStringCellValue());
        assertEquals("Pièce", sheet.getRow(0).getCell(2).getStringCellValue());
        assertEquals("Date", sheet.getRow(0).getCell(3).getStringCellValue());
        assertEquals("Journal", sheet.getRow(0).getCell(4).getStringCellValue());
        assertEquals("Contrepartie", sheet.getRow(0).getCell(5).getStringCellValue());
        assertEquals("Intitulé de la contrepartie", sheet.getRow(0).getCell(6).getStringCellValue());
        assertEquals("N° chèque", sheet.getRow(0).getCell(7).getStringCellValue());
        assertEquals("Libellé", sheet.getRow(0).getCell(8).getStringCellValue());
        assertEquals("Débit", sheet.getRow(0).getCell(9).getStringCellValue());
        assertEquals("Crédit", sheet.getRow(0).getCell(10).getStringCellValue());
        assertEquals("Solde (Calculé)", sheet.getRow(0).getCell(11).getStringCellValue());
        assertEquals("Vérification", sheet.getRow(0).getCell(12).getStringCellValue());
        assertEquals("Commentaire", sheet.getRow(0).getCell(13).getStringCellValue());
    }

    @Test
    void getCellsEnteteListeDesDepenses() {
        writeLine.getCellsEntete(sheet, NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES);
        assertEquals("Pièce", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Date", sheet.getRow(0).getCell(1).getStringCellValue());
        assertEquals("Libellé", sheet.getRow(0).getCell(2).getStringCellValue());
        assertEquals("Montant", sheet.getRow(0).getCell(3).getStringCellValue());
        assertEquals("Déduction", sheet.getRow(0).getCell(4).getStringCellValue());
        assertEquals("Récuperation", sheet.getRow(0).getCell(5).getStringCellValue());
        assertEquals("Commentaire", sheet.getRow(0).getCell(6).getStringCellValue());
    }

    @Test
    void getCellsEnteteEtatRapprochement() {
        writeLine.getCellsEnteteEtatRapprochement(sheet);
        assertEquals("Compte", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Intitulé du compte", sheet.getRow(0).getCell(1).getStringCellValue());
        assertEquals("Pièce", sheet.getRow(0).getCell(2).getStringCellValue());
        assertEquals("Date", sheet.getRow(0).getCell(3).getStringCellValue());
        assertEquals("Journal", sheet.getRow(0).getCell(4).getStringCellValue());
        assertEquals("Contrepartie", sheet.getRow(0).getCell(5).getStringCellValue());
        assertEquals("Intitulé de la contrepartie", sheet.getRow(0).getCell(6).getStringCellValue());
        assertEquals("N° chèque", sheet.getRow(0).getCell(7).getStringCellValue());
        assertEquals("Libellé", sheet.getRow(0).getCell(8).getStringCellValue());
        assertEquals("Débit", sheet.getRow(0).getCell(9).getStringCellValue());
        assertEquals("Crédit", sheet.getRow(0).getCell(10).getStringCellValue());
        assertEquals("----", sheet.getRow(0).getCell(11).getStringCellValue());
        assertEquals("Mois du relevé", sheet.getRow(0).getCell(12).getStringCellValue());
        assertEquals("Compte", sheet.getRow(0).getCell(13).getStringCellValue());
        assertEquals("Intitulé du compte", sheet.getRow(0).getCell(14).getStringCellValue());
        assertEquals("Date de l'opération", sheet.getRow(0).getCell(15).getStringCellValue());
        assertEquals("Date de valeur", sheet.getRow(0).getCell(16).getStringCellValue());
        assertEquals("Libellé", sheet.getRow(0).getCell(17).getStringCellValue());
        assertEquals("Débit", sheet.getRow(0).getCell(18).getStringCellValue());
        assertEquals("Crédit", sheet.getRow(0).getCell(19).getStringCellValue());
        assertEquals("Commentaire", sheet.getRow(0).getCell(20).getStringCellValue());
    }
}