package life.light.extract.info;

import life.light.type.InfoGrandLivre;
import life.light.type.TypeAccount;
import life.light.write.WriteFile;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Account {

    private static final Logger LOGGER = LogManager.getLogger();

    public Map<String, TypeAccount> getAccounts(String fileName, InfoGrandLivre infoGrandLivre, String codeCondominium) {
        Map<String, TypeAccount> accounts = new HashMap<>();
        Ledger ledger = new Ledger();
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (ledger.isAcccount(line, infoGrandLivre.postalCode(), codeCondominium)) {
                    TypeAccount account = ledger.account(line);
                    accounts.put(account.account(), account);
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Il y a {} comptes dans le grandlivre", accounts.size());

        String nameFile = "." + File.separator + "resultat" + File.separator + "Plan comptable de " + infoGrandLivre.syndicName();
        WriteFile writeFile = new WriteFile();
        writeFile.writeFileCSVAccounts(accounts, nameFile + ".csv");
        writeFile.writeFileExcelAccounts(accounts, nameFile + ".xlsx");

        return accounts;
    }
}