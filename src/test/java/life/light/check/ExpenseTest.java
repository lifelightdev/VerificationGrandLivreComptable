package life.light.check;

import life.light.FileOfTest;
import life.light.type.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.time.LocalDate;
import java.util.List;
import java.util.TreeMap;

import static life.light.FileOfTest.TEMP_TEST_DIR;
import static org.junit.jupiter.api.Assertions.*;

public class ExpenseTest {
    // Extract
    private Expense expense;
    private String testFilePath;

    @BeforeEach
    void setUp() {
        expense = new Expense();
        FileOfTest fileOfTest = new FileOfTest();
        testFilePath = fileOfTest.createExpenseListFile();
    }

    @Test
    void getExtractListOfExpense() {
        LineOfExpense[] result = expense.getListOfExpense(testFilePath);
        assertNotNull(result);
    }


    @Test
    void getCheckTotalNatureOK() {
        LineOfExpense[] listOfExpense = new LineOfExpense[1];
        List<String> result = expense.checkTotal(listOfExpense);
        assertTrue(result.isEmpty());
    }

    @Test
    void getCheckTotalNatureAmountKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "10", "10.00", "5.00", "2.00", TypeOfExpense.Nature);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour le code nature 010 :\s
                Le total du montant n'est pas correcte. Le total écrit est 10.00 alors que le calcul donne 100.0
                """, result.getFirst());
    }

    @Test
    void getCheckTotalNatureDeductionKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "10", "100.00", "50.00", "2.00", TypeOfExpense.Nature);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour le code nature 010 :\s
                Le total de la déduction n'est pas correcte. Le total écrit est 50.00 alors que le calcul donne 5.0
                """, result.getFirst());
    }

    @Test
    void getCheckTotalNatureRecoveryKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "20.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "10", "100.00", "5.00", "2.00", TypeOfExpense.Nature);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour le code nature 010 :\s
                Le total de la récuperation n'est pas correcte. Le total écrit est 2.00 alors que le calcul donne 20.0
                """, result.getFirst());
    }

    @Test
    void getCheckTotalNatureKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "10.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "10", "10.00", "15.00", "2.00", TypeOfExpense.Nature);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour le code nature 010 :\s
                Le total du montant n'est pas correcte. Le total écrit est 10.00 alors que le calcul donne 100.0
                Le total de la déduction n'est pas correcte. Le total écrit est 15.00 alors que le calcul donne 5.0
                Le total de la récuperation n'est pas correcte. Le total écrit est 2.00 alors que le calcul donne 12.0
                """, result.getFirst());
    }

    @Test
    void getCheckTotalKeyOK() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertTrue(result.isEmpty());
    }

    @Test
    void getCheckTotalKeyKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "10.00", "50.00", "12.00", TypeOfExpense.Key);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour le code clé 001 :\s
                Le total du montant n'est pas correcte. Le total écrit est 10.00 alors que le calcul donne 100.0
                Le total de la déduction n'est pas correcte. Le total écrit est 50.00 alors que le calcul donne 5.0
                Le total de la récuperation n'est pas correcte. Le total écrit est 12.00 alors que le calcul donne 2.0
                """, result.getFirst());
    }

    @Test
    void getCheckTotalBuldingOK() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBulding = new LineOfExpenseTotal("000", "Total de l'immeuble", "000", "100.00", "5.00", "2.00", TypeOfExpense.Key);


        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey,
                totalBulding
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertTrue(result.isEmpty());
    }

    @Test
    void getCheckTotalBuldingKO() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("1", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("2", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("3", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBulding = new LineOfExpenseTotal("000", "Total de l'immeuble", "000", "1100.00", "5.40", "2.10", TypeOfExpense.Building);


        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey,
                totalBulding
        };

        List<String> result = expense.checkTotal(listOfExpense);
        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals("""
                Pour l'immeuble :
                Le total du montant n'est pas correcte. Le total écrit est 1100.00 alors que le calcul donne 100.0
                Le total de la déduction n'est pas correcte. Le total écrit est 5.40 alors que le calcul donne 5.0
                Le total de la récuperation n'est pas correcte. Le total écrit est 2.10 alors que le calcul donne 2.0
                """, result.getFirst());
    }

    @Test
    void getCheckDocumentOK(){
        FileOfTest fileOfTest = new FileOfTest();
        try {
            fileOfTest.copyInvoiceFiles();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("Invoice123", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("Invoice234", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("Invoice345", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBulding = new LineOfExpenseTotal("000", "Total de l'immeuble", "000", "100.00", "5.00", "2.00", TypeOfExpense.Key);


        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey,
                totalBulding
        };

        TreeMap<String, String> listOfMissingDocuments = expense.checkDocument(listOfExpense, TEMP_TEST_DIR.toUri().toString());
        assertTrue(listOfMissingDocuments.isEmpty());
    }

    @Test
    @Disabled
    void getCheckDocumentKO(){
        FileOfTest fileOfTest = new FileOfTest();
        try {
            fileOfTest.deleteInvoiceFiles();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("Invoice123", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("Invoice234", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("Invoice345", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBulding = new LineOfExpenseTotal("000", "Total de l'immeuble", "000", "100.00", "5.00", "2.00", TypeOfExpense.Key);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey,
                totalBulding
        };

        TreeMap<String, String> listOfMissingDocuments = expense.checkDocument(listOfExpense, TEMP_TEST_DIR.toUri().toString());
        assertNotNull(listOfMissingDocuments);
        assertEquals(3, listOfMissingDocuments.size());
        assertEquals("Il manque la pièce Invoice123 pour le libellé : Dépense 1", listOfMissingDocuments.get("Invoice123"));
        assertEquals("Il manque la pièce Invoice234 pour le libellé : Dépense 2", listOfMissingDocuments.get("Invoice234"));
        assertEquals("Il manque la pièce Invoice345 pour le libellé : Dépense 3", listOfMissingDocuments.get("Invoice345"));
    }


    @Test
    @Disabled
    void writeFileListOfExpense() {
        LineOfExpenseKey key = new LineOfExpenseKey("Clé", "001", "Libellé de la clé 1", TypeOfExpense.Key);
        LineOfExpenseKey nature = new LineOfExpenseKey("Nature", "010", "Libellé de la nature 10", TypeOfExpense.Nature);

        LocalDate date = LocalDate.of(2023, 1, 15);
        LineOfExpenseValue expense1 = new LineOfExpenseValue("Invoice123", date, "Dépense 1", "50.00", "0.00", "0.00", "010", "001");
        LineOfExpenseValue expense2 = new LineOfExpenseValue("Invoice234", date, "Dépense 2", "30.00", "5.00", "0.00", "010", "001");
        LineOfExpenseValue expense3 = new LineOfExpenseValue("Invoice345", date, "Dépense 3", "20.00", "0.00", "2.00", "010", "001");

        LineOfExpenseTotal totalNature = new LineOfExpenseTotal("010", "Total de la nature", "010", "100.00", "5.00", "2.00", TypeOfExpense.Nature);
        LineOfExpenseTotal totalKey = new LineOfExpenseTotal("001", "Total de la clé", "001", "100.00", "5.00", "2.00", TypeOfExpense.Key);
        LineOfExpenseTotal totalBulding = new LineOfExpenseTotal("000", "Total de l'immeuble", "000", "100.00", "5.00", "2.00", TypeOfExpense.Key);

        // Create array with all objects in the correct order
        LineOfExpense[] listOfExpense = {
                key,
                nature,
                expense1,
                expense2,
                expense3,
                totalNature,
                totalKey,
                totalBulding
        };

        FileOfTest fileOfTest = new FileOfTest();
        try {
            fileOfTest.copyInvoiceFiles();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        TreeMap<String, String> listOfDocuments = expense.checkDocument(listOfExpense, TEMP_TEST_DIR.toUri().toString());

        File file = new File("." + File.separator + "TEST_temp" + File.separator + "Liste_des_depenses.xlsx");
        file.delete();
        expense.writeFileListOfExpense(listOfExpense, listOfDocuments);
        assertTrue(new File(file.getAbsolutePath()).exists());

    }

}
