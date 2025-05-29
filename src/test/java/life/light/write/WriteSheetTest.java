package life.light.write;

import life.light.type.Line;
import life.light.type.TotalAccount;
import life.light.type.TotalBuilding;
import life.light.type.TypeAccount;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.TreeSet;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class WriteSheetTest {

    private WriteSheet writeSheet;
    private Workbook workbook;
    private Object[] grandLivres;
    private TreeSet<String> journals;
    private String pathDirectoryInvoice;

    @BeforeEach
    void setUp() {
        writeSheet = new WriteSheet();
        workbook = new XSSFWorkbook();
        // Create test data
        Line line1 = new Line("DOC001", "2023-01-15", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK123", 
                "Paiement travaux", "1000.00", "");
        Line line2 = new Line("DOC002", "2023-01-20", 
                new TypeAccount("401", "Fournisseur"), "JR", 
                new TypeAccount("512", "Banque"), "CHK124", 
                "Remboursement", "", "500.00");
        Line line3 = new Line("DOC003", "2023-01-25", 
                new TypeAccount("512", "Banque"), "BQ", 
                new TypeAccount("401", "Fournisseur"), "CHK125", 
                "Autre opération", "500.00", "");
        TotalAccount totalAccount1 = new TotalAccount("Total compte Banque", 
                new TypeAccount("512", "Banque"), "1500.00", "0.00");
        TotalAccount totalAccount2 = new TotalAccount("Total compte Fournisseur", 
                new TypeAccount("401", "Fournisseur"), "0.00", "500.00");
        TotalBuilding totalBuilding = new TotalBuilding("Total général", "1500.00", "500.00");
        grandLivres = new Object[]{line1, line2, line3, totalAccount1, totalAccount2, totalBuilding};
        journals = new TreeSet<>();
        journals.add("BQ");
        journals.add("JR");
        pathDirectoryInvoice = "";
    }

    @Test
    void writeJournals() {
        // Execute the method
        writeSheet.writeJournals(grandLivres, journals, pathDirectoryInvoice, workbook);
        // Verify that sheets were created for each journal
        assertEquals(2, workbook.getNumberOfSheets());
        assertEquals("BQ", workbook.getSheetAt(0).getSheetName());
        assertEquals("JR", workbook.getSheetAt(1).getSheetName());
        // Verify content of BQ sheet
        Sheet bqSheet = workbook.getSheet("BQ");
        // Check header row
        Row headerRow = bqSheet.getRow(0);
        assertNotNull(headerRow);
        assertEquals("Compte", headerRow.getCell(0).getStringCellValue());
        // Check first data row
        Row dataRow1 = bqSheet.getRow(1);
        assertNotNull(dataRow1);
        assertEquals(512.0, dataRow1.getCell(0).getNumericCellValue());
        assertEquals("Banque", dataRow1.getCell(1).getStringCellValue());
        assertEquals("DOC001", dataRow1.getCell(2).getStringCellValue());
        // Verify content of JR sheet
        Sheet jrSheet = workbook.getSheet("JR");
        // Check first data row
        Row jrDataRow1 = jrSheet.getRow(1);
        assertNotNull(jrDataRow1);
        assertEquals(401.0, jrDataRow1.getCell(0).getNumericCellValue());
        assertEquals("Fournisseur", jrDataRow1.getCell(1).getStringCellValue());
        assertEquals("DOC002", jrDataRow1.getCell(2).getStringCellValue());
    }
}