package life.light;

import life.light.type.TypeAccount;
import life.light.write.WriteFile;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

class WriteFileTest {
    @Test
    void writeFileCSVAccountsTest() {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("51220", new TypeAccount("51220", "Banque"));
        accounts.put("10500", new TypeAccount("10500", "Fond travaux"));
        String filename = "." + File.separator + "temp" + File.separator + "ListeDesCompteTEST.csv";
        WriteFile.writeFileCSVAccounts(accounts, filename);
        try (BufferedReader reader = new BufferedReader(new FileReader(filename))) {
            String line = reader.readLine();
            Assertions.assertEquals("Compte;Intitulé du compte;", line);
            line = reader.readLine();
            Assertions.assertEquals("10500 ; Fond travaux ; ", line);
            line = reader.readLine();
            Assertions.assertEquals("51220 ; Banque ; ", line);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Erreur lors de la lecture du fichier CSV : " + e.getMessage());
        }
    }

    @Test
    void writeFileExcelAccountsTest() {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("51200", new TypeAccount("51200", "Banque"));
        accounts.put("10500", new TypeAccount("10500", "Fond travaux"));
        String filename = "." + File.separator + "temp" + File.separator + "ListeDesCompteTEST.xlsx";
        WriteFile.writeFileExcelAccounts(accounts, filename);
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
            Assertions.assertEquals("10500", row.getCell(0).getStringCellValue());
            Assertions.assertEquals("Fond travaux", row.getCell(1).getStringCellValue());
            row = sheet.getRow(2);
            Assertions.assertEquals("51200", row.getCell(0).getStringCellValue());
            Assertions.assertEquals("Banque", row.getCell(1).getStringCellValue());
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Erreur lors de la lecture du fichier Excel : " + e.getMessage());
        }
    }
}