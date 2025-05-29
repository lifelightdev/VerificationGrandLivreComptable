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
    static List<String> accountsbank = List.of("51220", "51221");
    static String pathDirectoryInvoice = "D:\\Le Nidor\\2024\\FACTURES";
    static String pathDirectoryBank = "D:\\Le Nidor\\2024\\BANQUE";
    static String pathFileLeger = "";
    static String pathFileListOfExpenses = "";
    static String codeCondominium = "";
    static int accountingYear = 2024;

    public static void main(String[] args) {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());

        if (args.length == 7) {
            codeCondominium = args[0];
            pathFileLeger = args[1];
            pathDirectoryBank = args[2];
            pathDirectoryInvoice = args[3];
            accountsbank = List.of(args[4]);
            pathFileListOfExpenses = args[5];
            accountingYear = Integer.parseInt(args[6]);
        }

        Expense expense = new Expense();
        int accountingYearOfListOfExpenses = expense.getAccountingYear(pathFileListOfExpenses);
        if (accountingYearOfListOfExpenses != accountingYear) {
            LOGGER.error("L'année comptable n'est pas cohérante ({},{})", accountingYearOfListOfExpenses, accountingYear);
        }
        Object[] listOfExpense = expense.getList(pathFileListOfExpenses);
        String path = "." + File.separator + "resultat" + File.separator + "Liste des dépenses.xlsx";
        WriteFile writeFile = new WriteFile();
        writeFile.writeFileExcelListeDesDepenses(listOfExpense, path, pathDirectoryInvoice);

        Ledger ledger = new Ledger(codeCondominium);
        InfoGrandLivre infoGrandLivre = ledger.getInfoGrandLivre(pathFileLeger);
        if (infoGrandLivre.stopDate().getYear() != accountingYear) {
            LOGGER.error("L'année comptable n'est pas cohérante ({},{})", infoGrandLivre.stopDate(), accountingYear);
        }

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts(pathFileLeger, codeCondominium);
        account.writeFilesAccounts(accounts, infoGrandLivre, "resultat");

        List<Line> lineBankInGrandLivre = ledger.getInfoBankGrandLivre(infoGrandLivre, accounts, pathFileLeger,
                pathDirectoryInvoice, accountsbank);

        // Récupération des relevés bancaire
        Bank bank = new Bank();
        List<BankLine> bankLines = bank.getBankLines(accounts, pathDirectoryBank, accountsbank,
                accountingYear);

        String nameFile = "." + File.separator + "resultat" + File.separator + "Etat de rapprochement.xlsx";
        writeFile.writeFileExcelEtatRaprochement(lineBankInGrandLivre, nameFile, bankLines);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}