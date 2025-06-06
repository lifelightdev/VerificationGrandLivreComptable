package life.light.extract.info;

import life.light.Constant;
import life.light.type.InfoGrandLivre;
import life.light.type.TypeAccount;
import life.light.write.WriteFile;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static life.light.Constant.*;

public class Account {

    private final Constant constant = new Constant();

    public Map<String, TypeAccount> getAccounts(String fileName, String codeCondominium) {
        Map<String, TypeAccount> accounts = new HashMap<>();
        Ledger ledger = new Ledger(codeCondominium);
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            accounts = reader.lines()
                    .filter(ledger::isAccount)
                    .map(ledger::account)
                    .collect(HashMap::new,
                            (map, account) -> map.put(account.account(), account),
                            HashMap::putAll);
        } catch (IOException e) {
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
        }
        return accounts;
    }

    public void writeFilesAccounts(Map<String, TypeAccount> accounts, InfoGrandLivre infoGrandLivre, String outputPath) {
        String nameFile = "." + File.separator + outputPath + File.separator + ACCOUNTING_PLAN + infoGrandLivre.syndicName();
        WriteFile writeFile = new WriteFile(PATH);
        writeFile.writeFileCSVAccounts(accounts, nameFile + CSV);
        writeFile.writeFileExcelAccounts(accounts, nameFile + XLSX);
    }
}