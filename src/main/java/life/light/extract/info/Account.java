package life.light.extract.info;

import life.light.Constant;
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

    public Map<String, TypeAccount> getAccounts() {
        String fileName = "Compte.csv";
        Map<String, TypeAccount> accounts = new HashMap<>();
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            for (String line = reader.readLine(); line != null; line = reader.readLine()) {
                String[] fields = line.split(";");
                if (!fields[0].equals("compte")) {
                    TypeAccount typeAccount = new TypeAccount(fields[0], fields[1]);
                    accounts.put(fields[0], typeAccount);
                }
            }
        } catch (IOException e) {
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
        }
        return accounts;
    }

    public void writeFilesAccounts(Map<String, TypeAccount> accounts, String outputPath) {
        String nameFile = "." + File.separator + outputPath + File.separator + ACCOUNTING_PLAN;
        WriteFile writeFile = new WriteFile(PATH);
        writeFile.writeFileCSVAccounts(accounts, nameFile + CSV);
        writeFile.writeFileExcelAccounts(accounts, nameFile + XLSX);
    }
}