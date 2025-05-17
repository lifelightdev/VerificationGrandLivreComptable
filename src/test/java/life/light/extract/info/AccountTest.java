package life.light.extract.info;

import life.light.type.TypeAccount;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Assertions;

import java.util.Map;

class AccountTest {

    @Test
    void getAccounts() {
        String fileName = ".\\temp\\grand_livre_test.txt";
        Map<String, TypeAccount> accounts = Account.getAccounts(fileName, "75000", "001" );
        Assertions.assertEquals(1, accounts.size());
        TypeAccount account = accounts.get("40100-0001");
        Assertions.assertEquals("40100-0001", account.account());
        Assertions.assertEquals("NOM du compte", account.label());
    }
}