package life.light.extract.info;

import life.light.type.TypeAccount;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class OutilInfoTest {

    private OutilInfo outilInfo;
    private Map<String, TypeAccount> accounts;

    @BeforeEach
    void setUp() {
        outilInfo = new OutilInfo();
        accounts = new HashMap<>();
        accounts.put("12345", new TypeAccount("12345", "Test Account"));
        accounts.put("40800", new TypeAccount("40800", "Fournisseurs"));
        accounts.put("42100", new TypeAccount("42100", "Personnel"));
        accounts.put("43100", new TypeAccount("43100", "Sécurité sociale"));
        accounts.put("43200", new TypeAccount("43200", "Autres organismes sociaux"));
        accounts.put("45000-123", new TypeAccount("45000-123", "Copropriétaire 123"));
        accounts.put("45000-456", new TypeAccount("45000-456", "Copropriétaire 456"));
        accounts.put("45000", new TypeAccount("45000", "Compte de tous les copropriétaires"));
    }

    @Test
    void testGetIndexOfNextWords() {
        // Given
        String[] words = {"word1", "", "word2"};
        // When & Then
        // The method only increments if the current word is empty
        assertEquals(0, outilInfo.getIndexOfNextWords(words, 0)); // word1 is not empty, so index stays the same
        assertEquals(2, outilInfo.getIndexOfNextWords(words, 1)); // "" is empty, so index is incremented
    }

    @Test
    void testGetWords() {
        // Given
        String line = "This is a test line";
        // When
        String[] words = outilInfo.getWords(line);
        // Then
        assertEquals(5, words.length);
        assertEquals("This", words[0]);
        assertEquals("is", words[1]);
        assertEquals("a", words[2]);
        assertEquals("test", words[3]);
        assertEquals("line", words[4]);
    }

    @Test
    void testIsAmount() {
        // Given & When & Then
        assertTrue(outilInfo.isAmount("100.00"));
        assertFalse(outilInfo.isAmount("100.00€"));
    }

    @Test
    void testGetNumberOfAmountsOn() {
        // Given
        String line = "Amount 100.00€ and another 200.00€";
        // When
        long count = outilInfo.getNumberOfAmountsOn(line);
        // Then
        assertEquals(2, count);
    }

    @Test
    void testGetAccount() {
        // Test regular account
        String[] words = {"12345"};
        TypeAccount account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("12345", account.account());
        // Test account with dash
        words = new String[]{"45000-123"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("45000-123", account.account());
        // Test co-owner account
        words = new String[]{"45000-456"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("45000-456", account.account());
        // Test co-owner account direct
        words = new String[]{"45000"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("45000", account.account());
        // Test special accounts
        words = new String[]{"40800123"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("40800", account.account());
        words = new String[]{"42100123"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("42100", account.account());
        words = new String[]{"43100123"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("43100", account.account());
        words = new String[]{"43200123"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNotNull(account);
        assertEquals("43200", account.account());
        // Test 9-digit account
        words = new String[]{"123456789"};
        account = outilInfo.getAccount(accounts, words, 0);
        assertNull(account); // Not in our test accounts map
    }

    @Test
    void testFixedSpacesBeforeEuroSign() {
        // Given
        String line = "Amount 100.00 €";
        // When
        String result = outilInfo.fixedSpacesBeforeEuroSign(line);
        // Then
        assertEquals("Amount 100.00€", result);
    }

    @Test
    void testSplittingLineIntoWordTable() {
        // Test normal line
        String line = "This is a test";
        String[] result = outilInfo.splittingLineIntoWordTable(line);
        assertEquals(4, result.length);
        assertEquals("This", result[0]);
        assertEquals("is", result[1]);
        assertEquals("a", result[2]);
        assertEquals("test", result[3]);
        // Test line with multiple spaces
        line = "This  is  a  test";
        result = outilInfo.splittingLineIntoWordTable(line);
        assertEquals(7, result.length);
        assertEquals("This", result[0]);
        assertEquals(" ", result[1]);
        assertEquals("is", result[2]);
        assertEquals(" ", result[3]);
        assertEquals("a", result[4]);
        assertEquals(" ", result[5]);
        assertEquals("test", result[6]);
        // Test line ending with space
        line = "This is a test ";
        result = outilInfo.splittingLineIntoWordTable(line);
        assertEquals(5, result.length);
        assertEquals("This", result[0]);
        assertEquals("is", result[1]);
        assertEquals("a", result[2]);
        assertEquals("test", result[3]);
        assertEquals(" ", result[4]);
    }

    @Test
    void testGetDocument() {
        // Test normal document
        String[] words = {"DOC123"};
        String result = outilInfo.getDocument(words, 0);
        assertEquals("OC123", result); // The method removes the first character if document length is 6
        // Test 6-character document
        words = new String[]{"D12345"};
        result = outilInfo.getDocument(words, 0);
        assertEquals("12345", result); // The method removes the first character if document length is 6
        // Test date (should return empty)
        words = new String[]{"01/01/2023"};
        result = outilInfo.getDocument(words, 0);
        assertEquals("", result);
        // Test phone number (should return empty)
        words = new String[]{"0123456789"};
        result = outilInfo.getDocument(words, 0);
        assertEquals("", result);
    }

    @Test
    void testRemovesStrayCharactersInLine() {
        // Given
        String line = "This | is / a Reportde test ' — . =— _ - = l";
        // When
        String result = outilInfo.removesStrayCharactersInLine(line);
        // Then
        // The method doesn't remove the single quote character and there's a space at the end
        assertEquals("This is a Report de test ' ", result);
    }

    @Test
    void testFindDateIn() {
        // Test dd/MM/yyyy format
        String line = "Date: 01/01/2023";
        String result = outilInfo.findDateIn(line);
        assertEquals("01/01/2023", result);
        // Test dd.MM.yy format
        line = "Date: 01.01.23";
        result = outilInfo.findDateIn(line);
        assertEquals("01.01.23", result);
        // Test no date
        line = "No date here";
        result = outilInfo.findDateIn(line);
        assertEquals("", result);
    }

    @Test
    void testGetNumberOfLineInFile(@TempDir Path tempDir) throws IOException {
        // Given
        File testFile = tempDir.resolve("test.txt").toFile();
        try (FileWriter writer = new FileWriter(testFile)) {
            writer.write("Line 1\nLine 2\n\nLine 3");
        }
        // When
        int count = outilInfo.getNumberOfLineInFile(testFile.getAbsolutePath());
        // Then
        assertEquals(3, count);
    }
}