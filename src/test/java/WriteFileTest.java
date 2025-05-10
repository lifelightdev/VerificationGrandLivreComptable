import life.light.Account;
import life.light.WriteFile;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class WriteFileTest {
    @Test
    public void writeFileCSVAccountsTest() {
        Map<String, Account> accounts = new HashMap<>();
        accounts.put("51220", new Account("51220", "Banque"));
        accounts.put("10500", new Account("10500", "Fond travaux"));
        String filename = "." + File.separator + "temp" + File.separator + "ListeDesCompteTEST.csv";
        WriteFile.writeFileCSVAccounts(accounts, filename);
        try (BufferedReader reader = new BufferedReader(new FileReader(filename))) {
            String line = reader.readLine();
            Assertions.assertEquals(line, "Compte;Intitulé du compte;");
            line = reader.readLine();
            Assertions.assertEquals(line, "10500 ; Fond travaux ; ");
            line = reader.readLine();
            Assertions.assertEquals(line, "51220 ; Banque ; ");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}