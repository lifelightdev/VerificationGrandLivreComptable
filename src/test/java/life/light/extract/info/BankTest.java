package life.light.extract.info;

import life.light.FileOfTest;
import life.light.type.BankLine;
import life.light.type.TypeAccount;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static life.light.FileOfTest.tempTestDir;
import static org.junit.jupiter.api.Assertions.assertEquals;

class BankTest {

    Map<String, TypeAccount> accounts = new HashMap<>();
    TypeAccount accountBank1 = new TypeAccount("51220", "Banque 1");
    TypeAccount accountBank2 = new TypeAccount("51221", "Banque 2");
    Bank bank = new Bank();

    @BeforeEach
    void setUp() {
        accounts.put(accountBank1.account(), accountBank1);
        accounts.put(accountBank2.account(), accountBank2);
        FileOfTest fileOfTest = new FileOfTest();
        try {
            fileOfTest.copyBankFiles(accountBank1.account(), accountBank2.account());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void getBankLines() {
        String pathsDirectoryBank = tempTestDir + File.separator + "bank" + File.separator;
        List<String> accountsbank = List.of(accountBank2.account());
        List<BankLine> bankLines = bank.getBankLines(accounts, pathsDirectoryBank, accountsbank, 2024);
        assertEquals(2, bankLines.size());
        LocalDate operationDate = LocalDate.parse("2024-07-15");
        assertEquals(operationDate, bankLines.getLast().operationDate());
        LocalDate valueDate = LocalDate.parse("2024-07-16");
        assertEquals(valueDate, bankLines.getLast().valueDate());
        String label = "REM CHQ N° 000000001";
        assertEquals(label, bankLines.getLast().label());
        double debit = 0;
        assertEquals(debit, bankLines.getLast().debit());
        double credit = 100.00;
        assertEquals(credit, bankLines.getLast().credit());
        assertEquals(accountBank2.account(), bankLines.getLast().account().account());
    }

    @Test
    void getBankMultiLines() {
        String pathsDirectoryBank = tempTestDir + File.separator + "bank" + File.separator;
        List<String> accountsbank = List.of(accountBank2.account());
        List<BankLine> bankLines = bank.getBankLines(accounts, pathsDirectoryBank, accountsbank, 2024);
        assertEquals(2, bankLines.size());
        LocalDate operationDate = LocalDate.parse("2024-08-17");
        assertEquals(operationDate, bankLines.getFirst().operationDate());
        LocalDate valueDate = LocalDate.parse("2024-08-18");
        assertEquals(valueDate, bankLines.getFirst().valueDate());
        String label = "REM CHQ N° 000000001 Avec un text sur plusieur ligne 3T2024";
        assertEquals(label, bankLines.getFirst().label());
        double debit = 0;
        assertEquals(debit, bankLines.getFirst().debit());
        double credit = 100.00;
        assertEquals(credit, bankLines.getFirst().credit());
        assertEquals(accountBank2.account(), bankLines.getFirst().account().account());
    }

    @Test
    void getMultiBank() {
        String pathsDirectoryBank = tempTestDir + File.separator + "bank" + File.separator;
        List<String> accountsbank = List.of(accountBank1.account(), accountBank2.account());
        List<BankLine> bankLines = bank.getBankLines(accounts, pathsDirectoryBank, accountsbank, 2024);
        assertEquals(6, bankLines.size());
        LocalDate operationDate = LocalDate.parse("2024-11-01");
        assertEquals(operationDate, bankLines.getFirst().operationDate());
        LocalDate valueDate = LocalDate.parse("2024-11-01");
        assertEquals(valueDate, bankLines.getFirst().valueDate());
        String label = "VIR INST M. TINTIN LIBELLE:Tintin ref : 1.018.0031 REF.CLIENT-NOTPROVIDED";
        assertEquals(label, bankLines.getFirst().label());
        double debit = 0;
        assertEquals(debit, bankLines.getFirst().debit());
        double credit = 1500.88;
        assertEquals(credit, bankLines.getFirst().credit());
        assertEquals(accountBank1.account(), bankLines.getFirst().account().account());
    }

    @Test
    void getAccountsBank() {
        List<TypeAccount> accountsBank = Bank.getBanks(accounts);
        assertEquals(2, accountsBank.size());
        assertEquals(accountBank2, accountsBank.get(0));
        assertEquals(accountBank1, accountsBank.get(1));
    }
}