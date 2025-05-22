package life.light.extract.info;

import life.light.FileOfTest;
import life.light.type.InfoGrandLivre;
import life.light.type.TypeAccount;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.util.Map;

class AccountTest {

    String nameFileTestLedger = "";
    Account account = new Account();

    @BeforeEach
    void setUp() {
        FileOfTest fileOfTest = new FileOfTest();
        try {
            nameFileTestLedger = fileOfTest.createMinimalLedgerFile();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void getAccounts() {
        InfoGrandLivre infoGrandLivre = new InfoGrandLivre("Name", "31/12/2024", LocalDate.of(2024, 12, 31), "75000");
        Map<String, TypeAccount> accounts = account.getAccounts(nameFileTestLedger, infoGrandLivre, "001");
        Assertions.assertEquals(1, accounts.size());
        TypeAccount account = accounts.get("40100-0001");
        Assertions.assertEquals("40100-0001", account.account());
        Assertions.assertEquals("NOM du compte", account.label());
    }
}