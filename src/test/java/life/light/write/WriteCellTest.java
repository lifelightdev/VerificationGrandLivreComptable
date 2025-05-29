package life.light.write;

import life.light.type.CellValues;
import life.light.type.Line;
import life.light.type.TypeAccount;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static life.light.write.WriteOutil.*;
import static org.junit.jupiter.api.Assertions.*;

class WriteCellTest {

    private WriteCell writeCell;
    private Workbook workbook;
    private Sheet sheet;
    private WriteCellStyle writeCellStyle;

    @BeforeEach
    void setUp() {
        writeCell = new WriteCell();
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Test Sheet");
        writeCellStyle = new WriteCellStyle();
    }

    @Test
    void addCell_withNumericValue() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        writeCell.addCell(row, 0, "123.45", style, "test line", "test name", "test place");
        assertEquals(123.45, row.getCell(0).getNumericCellValue());
        assertEquals(style, row.getCell(0).getCellStyle());
    }

    @Test
    void addCell_withTextValue() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        writeCell.addCell(row, 0, "test value", style, "test line", "test name", "test place");
        assertEquals("test value", row.getCell(0).getStringCellValue());
        assertEquals(style, row.getCell(0).getCellStyle());
    }

    @Test
    void addCell_withEmptyValue() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        writeCell.addCell(row, 0, "", style, "test line", "test name", "test place");
        // Cell should be created but empty
        assertNotNull(row.getCell(0));
        assertEquals(style, row.getCell(0).getCellStyle());
    }

    @Test
    void addCells() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        List<CellValues> values = new ArrayList<>();
        values.add(new CellValues(0, "value1", style, "line1", "name1"));
        values.add(new CellValues(1, "123.45", style, "line2", "name2"));
        writeCell.addCells(row, values, "test place");
        assertEquals("value1", row.getCell(0).getStringCellValue());
        assertEquals(123.45, row.getCell(1).getNumericCellValue());
    }

    @Test
    void addCellEmpty() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        writeCell.addCellEmpty(0, 3, row, style);
        // Check that cells were created with the style
        assertNotNull(row.getCell(0));
        assertNotNull(row.getCell(1));
        assertNotNull(row.getCell(2));
        assertEquals(style, row.getCell(0).getCellStyle());
        assertEquals(style, row.getCell(1).getCellStyle());
        assertEquals(style, row.getCell(2).getCellStyle());
    }

    @Test
    void addSoldeCell_normalLine() {
        Row row = sheet.createRow(2); // Not first row
        CellStyle style = workbook.createCellStyle();
        // Create previous row with solde cell
        Row prevRow = sheet.createRow(1);
        Cell prevSoldeCell = prevRow.createCell(ID_BALANCE_OF_LEDGER);
        // Create debit and credit cells
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        Cell soldeCell = writeCell.addSoldeCell(row, debitCell, creditCell, style, ID_BALANCE_OF_LEDGER, false, false);
        // Formula should be: previous_solde + debit - credit
        String expectedFormula = prevSoldeCell.getAddress() + "+" + debitCell.getAddress() + "-" + creditCell.getAddress();
        assertEquals(expectedFormula, soldeCell.getCellFormula());
        assertEquals(style, soldeCell.getCellStyle());
    }

    @Test
    void addSoldeCell_firstLine() {
        Row row = sheet.createRow(1); // First row
        CellStyle style = workbook.createCellStyle();
        // Create debit and credit cells
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        Cell soldeCell = writeCell.addSoldeCell(row, debitCell, creditCell, style, ID_BALANCE_OF_LEDGER, false, false);
        // Formula for first line should be: credit - debit
        String expectedFormula = creditCell.getAddress() + "-" + debitCell.getAddress();
        assertEquals(expectedFormula, soldeCell.getCellFormula());
        assertEquals(style, soldeCell.getCellStyle());
    }

    @Test
    void addSoldeCell_reportLine() {
        Row row = sheet.createRow(2);
        CellStyle style = workbook.createCellStyle();
        // Create debit and credit cells
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        Cell soldeCell = writeCell.addSoldeCell(row, debitCell, creditCell, style, ID_BALANCE_OF_LEDGER, true, false);
        // Formula for report line should be: credit - debit
        String expectedFormula = creditCell.getAddress() + "-" + debitCell.getAddress();
        assertEquals(expectedFormula, soldeCell.getCellFormula());
        assertEquals(style, soldeCell.getCellStyle());
    }

    @Test
    void addSoldeCell_totalLine() {
        Row row = sheet.createRow(2);
        CellStyle style = workbook.createCellStyle();
        // Create debit and credit cells
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        Cell soldeCell = writeCell.addSoldeCell(row, debitCell, creditCell, style, ID_BALANCE_OF_LEDGER, false, true);
        // Formula for total line should be: debit - credit
        String expectedFormula = debitCell.getAddress() + "-" + creditCell.getAddress();
        assertEquals(expectedFormula, soldeCell.getCellFormula());
        assertEquals(style, soldeCell.getCellStyle());
    }

    @Test
    void addCreditCell() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        Line line = new Line("DOC001", "2023-01-15", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK123", 
                "Test operation", "0.00", "100.50");
        Cell creditCell = writeCell.addCreditCell(line, row, style);
        assertEquals(100.50, creditCell.getNumericCellValue());
        assertEquals(style, creditCell.getCellStyle());
    }

    @Test
    void addCreditCell_nonNumeric() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        Line line = new Line("DOC001", "2023-01-15", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK123", 
                "Test operation", "0.00", "");
        Cell creditCell = writeCell.addCreditCell(line, row, style);
        assertEquals(0.0, creditCell.getNumericCellValue());
        assertEquals(style, creditCell.getCellStyle());
    }

    @Test
    void addCellCreditOfTotalAccountInLedger() {
        Row row = sheet.createRow(3);
        CellStyle style = writeCellStyle.getCellStyleAmount(workbook, false);
        // Create previous rows with credit values
        Row row1 = sheet.createRow(1);
        Cell credit1 = row1.createCell(ID_CREDIT_OF_LEDGER);
        credit1.setCellValue(100.0);
        Row row2 = sheet.createRow(2);
        Cell credit2 = row2.createCell(ID_CREDIT_OF_LEDGER);
        credit2.setCellValue(200.0);
        Cell creditTotalCell = writeCell.addCellCreditOfTotalAccountInLedger(row, 0, style);
        // Formula should sum from row 1 to row 2
        String expectedFormula = "SUM(" + credit1.getAddress().formatAsString().replace("$", "") + 
                ":" + credit2.getAddress().formatAsString().replace("$", "") + ")";
        assertEquals(expectedFormula, creditTotalCell.getCellFormula());
        assertEquals(style, creditTotalCell.getCellStyle());
    }

    @Test
    void addDebitCell() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        Line line = new Line("DOC001", "2023-01-15", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK123", 
                "Test operation", "100.50", "0.00");
        Cell debitCell = writeCell.addDebitCell(line, row, style);
        assertEquals(100.50, debitCell.getNumericCellValue());
        assertEquals(style, debitCell.getCellStyle());
    }

    @Test
    void addDebitCell_nonNumeric() {
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        Line line = new Line("DOC001", "2023-01-15", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK123", 
                "Test operation", "", "0.00");
        Cell debitCell = writeCell.addDebitCell(line, row, style);
        assertEquals(0.0, debitCell.getNumericCellValue());
        assertEquals(style, debitCell.getCellStyle());
    }

    @Test
    void addCellDebitOfTotalAccountInLedger() {
        Row row = sheet.createRow(3);
        CellStyle style = writeCellStyle.getCellStyleAmount(workbook, false);
        // Create previous rows with debit values
        Row row1 = sheet.createRow(1);
        Cell debit1 = row1.createCell(ID_DEBIT_OF_LEDGER);
        debit1.setCellValue(100.0);
        Row row2 = sheet.createRow(2);
        Cell debit2 = row2.createCell(ID_DEBIT_OF_LEDGER);
        debit2.setCellValue(200.0);
        Cell debitTotalCell = writeCell.addCellDebitOfTotalAccountInLedger(row, 0, style);
        // Formula should sum from row 1 to row 2
        String expectedFormula = "SUM(" + debit1.getAddress().formatAsString().replace("$", "") + 
                ":" + debit2.getAddress().formatAsString().replace("$", "") + ")";
        assertEquals(expectedFormula, debitTotalCell.getCellFormula());
        assertEquals(style, debitTotalCell.getCellStyle());
    }

    @Test
    void addCellAmountOfTotalBuildingInLedger() {
        Row row = sheet.createRow(5);
        CellStyle style = writeCellStyle.getCellStyleTotalAmount(workbook);
        // Create rows with total values at specific row numbers
        Row row2 = sheet.createRow(2);
        Cell total1 = row2.createCell(ID_DEBIT_OF_LEDGER);
        total1.setCellValue(100.0);
        Row row4 = sheet.createRow(4);
        Cell total2 = row4.createCell(ID_DEBIT_OF_LEDGER);
        total2.setCellValue(200.0);
        List<Integer> lineTotals = new ArrayList<>();
        lineTotals.add(2);
        lineTotals.add(4);
        Cell totalCell = writeCell.addCellAmountOfTotalBuildingInLedger(row, ID_DEBIT_OF_LEDGER, lineTotals, style);
        // Formula should be J2+J4 (assuming ID_DEBIT_OF_LEDGER is column J)
        String columnLetter = "J"; // This is the column for ID_DEBIT_OF_LEDGER (9)
        String expectedFormula = columnLetter + "2+" + columnLetter + "4";
        assertEquals(expectedFormula, totalCell.getCellFormula());
        assertEquals(style, totalCell.getCellStyle());
    }

    @Test
    void addCellTotalAmount() {
        Row row = sheet.createRow(3);
        CellStyle style = writeCellStyle.getCellStyleAmount(workbook, false);
        // Create previous rows with values
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue(100.0);
        Row row2 = sheet.createRow(2);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(200.0);
        Cell totalCell = writeCell.addCellTotalAmount(row, 0, 0, style);
        // Formula should sum from row 1 to row 2
        String expectedFormula = "SUM(" + cell1.getAddress().formatAsString().replace("$", "") + 
                ":" + cell2.getAddress().formatAsString().replace("$", "") + ")";
        assertEquals(expectedFormula, totalCell.getCellFormula());
        assertEquals(style, totalCell.getCellStyle());
    }
}