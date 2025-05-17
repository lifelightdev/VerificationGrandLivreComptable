package life.light.extract.info;

import life.light.type.TypeAccount;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Account {

    private static final Logger LOGGER = LogManager.getLogger();

    private Account() {    }

    public static Map<String, TypeAccount> getAccounts(String fileName, String postalCode, String codeCondominium ) {
        Map<String, TypeAccount> accounts = new HashMap<>();
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (Ledger.isAcccount(line, postalCode, codeCondominium)) {
                    TypeAccount account = Ledger.account(line);
                    accounts.put(account.account(), account);
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Il y a {} comptes dans le grandlivre", accounts.size());
        return accounts;
    }
}