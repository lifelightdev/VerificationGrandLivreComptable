package life.light.extract.info;

import life.light.FileOfTest;
import life.light.type.TypeAccount;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

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
        Map<String, TypeAccount> accounts = account.getAccounts(nameFileTestLedger, "001");
        Assertions.assertEquals(1, accounts.size());
        TypeAccount account = accounts.get("40100-0001");
        Assertions.assertEquals("40100-0001", account.account());
        Assertions.assertEquals("NOM du compte", account.label());
    }
}