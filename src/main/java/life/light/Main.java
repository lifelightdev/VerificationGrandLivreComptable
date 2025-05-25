package life.light;

import life.light.extract.info.Account;
import life.light.extract.info.Bank;
import life.light.extract.info.Expense;
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
import java.util.List;
import java.util.Map;

public class Main {

    private static final Logger LOGGER = LogManager.getLogger();
    static List<String> accountsbank = List.of();
    static String pathDirectoryInvoice = "";
    static String pathDirectoryBank = "";
    static String pathDirectoryLeger = "";
    static String pathDirectoryListOfExpenses = "";
    static String codeCondominium = "";

    public static void main(String[] args) {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());

        if (args.length == 6) {
            codeCondominium = args[0];
            pathDirectoryLeger = args[1];
            pathDirectoryBank = args[2];
            pathDirectoryInvoice = args[3];
            accountsbank = List.of(args[4]);
            pathDirectoryListOfExpenses = args[5];
        }

        Expense expense = new Expense();
        Object[] listOfExpense = expense.getList(pathDirectoryListOfExpenses);
        String path = "." + File.separator + "resultat" + File.separator + "Liste des dépenses.xlsx";
        WriteFile writeFile = new WriteFile();
        writeFile.writeFileExcelListeDesDepenses(listOfExpense, path, pathDirectoryInvoice);

        Ledger ledger = new Ledger(codeCondominium);
        InfoGrandLivre infoGrandLivre = ledger.getInfoGrandLivre(pathDirectoryLeger);

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts(pathDirectoryLeger, codeCondominium);
        account.writeFilesAccounts(accounts, infoGrandLivre, "resultat");

        List<Line> lineBankInGrandLivre = ledger.getInfoBankGrandLivre(infoGrandLivre, accounts, pathDirectoryLeger,
                pathDirectoryInvoice, accountsbank);

        // Récupération des relevés bancaire
        Bank bank = new Bank();
        List<BankLine> bankLines = bank.getBankLines(accounts, pathDirectoryBank, accountsbank,
                infoGrandLivre.stopDate().getYear());

        String nameFile = "." + File.separator + "resultat" + File.separator + "Etat de rapprochement.xlsx";
        writeFile.writeFileExcelEtatRaprochement(lineBankInGrandLivre, nameFile, bankLines);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}