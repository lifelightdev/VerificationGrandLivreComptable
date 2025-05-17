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
        accounts.put("512", new TypeAccount("512", "Banque"));
        List<String> pathsDirectoryBank = new ArrayList<>();
        pathsDirectoryBank.add(".\\temp\\bank\\512\\");
        List<BankLine> bankLines = Bank.getBankLines(accounts, pathsDirectoryBank);
        assertEquals(1, bankLines.size());
        LocalDate operationDate =  LocalDate.parse("2024-07-15");
        assertEquals(operationDate, bankLines.getFirst().operationDate());
        LocalDate valueDate =  LocalDate.parse("2024-07-16");
        assertEquals(valueDate, bankLines.getFirst().valueDate());
        String label = "REM CHQ N° 000000001";
        assertEquals(label, bankLines.getFirst().label());
        double debit = 0;
        assertEquals(debit, bankLines.getFirst().debit());
        double credit = 100.00;
        assertEquals(credit, bankLines.getFirst().credit());
        assertEquals("512", bankLines.getFirst().account().account());
    }
}