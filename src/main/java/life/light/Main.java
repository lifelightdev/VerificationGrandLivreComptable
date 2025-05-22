package life.light;

import life.light.extract.info.Account;
import life.light.extract.info.Bank;
import life.light.extract.info.Ledger;
import life.light.type.BankLine;
import life.light.type.InfoGrandLivre;
import life.light.type.Line;
import life.light.type.TypeAccount;
import life.light.write.WriteFile;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class Main {

    private static final Logger LOGGER = LogManager.getLogger();
    // TODO : à mettre dans les paramètres de lancement du programme
    public static final String BANK_1_ACCOUNT = "";
    public static final String BANK_2_ACCOUNT = "";
    static String pathDirectoryInvoice = "";
    static String pathDirectoryBank = "";
    static String pathDirectoryLeger = "";
    static String codeCondominium = "";

    public static void main(String[] args) {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());

        if (args.length == 4) {
            codeCondominium = args[0];
            pathDirectoryLeger = args[1];
            pathDirectoryBank = args[2];
            pathDirectoryInvoice = args[3];
        }

        Ledger ledger = new Ledger();
        InfoGrandLivre infoGrandLivre = ledger.getInfoGrandLivre(pathDirectoryLeger);

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts(pathDirectoryLeger, infoGrandLivre, codeCondominium);

        List<Line> lineBankInGrandLivre = ledger.getInfoBankGrandLivre(infoGrandLivre, accounts, pathDirectoryLeger, pathDirectoryInvoice);

        // Récupération des relevés bancaire
        List<String> pathsDirectoryBank = new ArrayList<>();
        pathsDirectoryBank.add(pathDirectoryBank + File.separator + BANK_2_ACCOUNT);
        pathsDirectoryBank.add(pathDirectoryBank + File.separator + BANK_1_ACCOUNT);
        Bank bank = new Bank();
        List<BankLine> bankLines = bank.getBankLines(accounts, pathsDirectoryBank, infoGrandLivre.stopDate().getYear());

        String nameFile = "." + File.separator + "temp" + File.separator + "Etat de rapprochement.xlsx";
        WriteFile writeFile = new WriteFile();
        writeFile.writeFileExcelEtatRaprochement(lineBankInGrandLivre, nameFile, bankLines);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}