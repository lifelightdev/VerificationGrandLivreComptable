package life.light.extract.info;

import life.light.type.BankLine;
import life.light.type.TypeAccount;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class BankTest {

    @Test
    void getBankLines() {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("5120", new TypeAccount("5120", "Banque 1"));
        List<String> pathsDirectoryBank = new ArrayList<>();
        pathsDirectoryBank.add(".\\temp\\bank\\5120\\");
        List<BankLine> bankLines = Bank.getBankLines(accounts, pathsDirectoryBank);
        assertEquals(2, bankLines.size());
        LocalDate operationDate =  LocalDate.parse("2024-07-15");
        assertEquals(operationDate, bankLines.getLast().operationDate());
        LocalDate valueDate =  LocalDate.parse("2024-07-16");
        assertEquals(valueDate, bankLines.getLast().valueDate());
        String label = "REM CHQ N° 000000001";
        assertEquals(label, bankLines.getLast().label());
        double debit = 0;
        assertEquals(debit, bankLines.getLast().debit());
        double credit = 100.00;
        assertEquals(credit, bankLines.getLast().credit());
        assertEquals("5120", bankLines.getLast().account().account());
    }

    @Test
    void getBankMultiLines() {
        Map<String, TypeAccount> accounts = new HashMap<>();
        accounts.put("5120", new TypeAccount("5120", "Banque 1"));
        List<String> pathsDirectoryBank = new ArrayList<>();
        pathsDirectoryBank.add(".\\temp\\bank\\5120\\");
        List<BankLine> bankLines = Bank.getBankLines(accounts, pathsDirectoryBank);
        assertEquals(2, bankLines.size());
        LocalDate operationDate =  LocalDate.parse("2024-08-17");
        assertEquals(operationDate, bankLines.getFirst().operationDate());
        LocalDate valueDate =  LocalDate.parse("2024-08-18");
        assertEquals(valueDate, bankLines.getFirst().valueDate());
        String label = "REM CHQ N° 000000001 Avec un text sur plusieur ligne 3T2024";
        assertEquals(label, bankLines.getFirst().label());
        double debit = 0;
        assertEquals(debit, bankLines.getFirst().debit());
        double credit = 100.00;
        assertEquals(credit, bankLines.getFirst().credit());
        assertEquals("5120", bankLines.getFirst().account().account());
    }
}