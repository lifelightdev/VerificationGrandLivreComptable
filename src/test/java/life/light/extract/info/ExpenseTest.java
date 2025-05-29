package life.light.extract.info;

import life.light.FileOfTest;
import life.light.type.LineOfExpense;
import life.light.type.LineOfExpenseKey;
import life.light.type.LineOfExpenseTotal;
import life.light.type.TypeOfExpense;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import static org.junit.jupiter.api.Assertions.*;

class ExpenseTest {

    private Expense expense;
    private String testFilePath;

    @BeforeEach
    void setUp() {
        expense = new Expense();
        FileOfTest fileOfTest = new FileOfTest();
        testFilePath = fileOfTest.createExpenseListFile();
    }

    @Test
    void getList() {
        // When
        Object[] result = expense.getList(testFilePath);

        // Then
        assertNotNull(result);
        assertTrue(result.length > 0);

        LineOfExpenseKey key = (LineOfExpenseKey) result[0];
        assertEquals("001", key.key());
        assertEquals("C'est la première clé", key.value());
        assertEquals(TypeOfExpense.Key, key.type());

        // Verify nature lines
        LineOfExpenseKey nature = (LineOfExpenseKey) result[1];
        assertEquals("61100", nature.key());
        assertEquals("Nettoyage", nature.value());
        assertEquals(TypeOfExpense.Nature, nature.type());

        // Verify expense lines
        LineOfExpense expenseLine = (LineOfExpense) result[2];
        assertEquals("33333", expenseLine.document());
        assertEquals(LocalDate.parse("01/01/2024", DateTimeFormatter.ofPattern("dd/MM/yyyy")), expenseLine.date());
        assertEquals("Facture de nettoyage", expenseLine.label());
        assertEquals("100.00", expenseLine.amount());
        assertEquals("0.00", expenseLine.deduction());
        assertEquals("0.00", expenseLine.recovery());

        // Verify expense lines
        LineOfExpense expenseLine1 = (LineOfExpense) result[7];
        assertEquals("38688", expenseLine1.document());
        assertEquals(LocalDate.parse("15/04/2024", DateTimeFormatter.ofPattern("dd/MM/yyyy")), expenseLine1.date());
        assertEquals("Remplacement", expenseLine1.label());
        assertEquals("1800.00", expenseLine1.amount());
        assertEquals("1800.00", expenseLine1.recovery());

        // Verify total lines
        LineOfExpenseTotal totalNature = (LineOfExpenseTotal) result[4];
        assertEquals("61100", totalNature.key());
        assertEquals("250.00", totalNature.amount());
        assertEquals("0.00", totalNature.deduction());
        assertEquals("0.00", totalNature.recovery());
        assertEquals(TypeOfExpense.Nature, totalNature.type());

        LineOfExpenseTotal totalKey = (LineOfExpenseTotal) result[9];
        assertEquals("001", totalKey.key());
        assertEquals("450.00", totalKey.amount());
        assertEquals("0.00", totalKey.deduction());
        assertEquals("0.00", totalKey.recovery());
        assertEquals(TypeOfExpense.Key, totalKey.type());


        LineOfExpenseTotal totalBuilding = (LineOfExpenseTotal) result[10];
        assertEquals("570.00", totalBuilding.amount());
        assertEquals("0.00", totalBuilding.deduction());
        assertEquals("0.00", totalBuilding.recovery());
        assertEquals(TypeOfExpense.Building, totalBuilding.type());
    }

    @Test
    void getAccountingYear() {
        int year = expense.getAccountingYear(testFilePath);
        assertEquals(2024, year);
    }
}