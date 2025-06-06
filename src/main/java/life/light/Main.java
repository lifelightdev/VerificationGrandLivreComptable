package life.light;

import life.light.check.Expense;
import life.light.extract.info.Account;
import life.light.extract.info.Bank;
import life.light.extract.info.ExpenseExtract;
import life.light.extract.info.Ledger;
import life.light.type.*;
import life.light.write.WriteFile;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import static life.light.Constant.PATH;

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

        ExpenseExtract expense = new ExpenseExtract();
        int accountingYearOfListOfExpenses = expense.getAccountingYear(pathFileListOfExpenses);
        if (accountingYearOfListOfExpenses != accountingYear) {
            LOGGER.error("L'année comptable n'est pas cohérente (dans la liste des dépense ; {}, passée en paramètre : {})", accountingYearOfListOfExpenses, accountingYear);
        }
        Expense expenseCheck = new Expense();
        LineOfExpense[] listOfExpense = expenseCheck.getListOfExpense(pathFileListOfExpenses);
        List<String> checkTotal = expenseCheck.checkTotal(listOfExpense);
        TreeMap<String, String> checkDocument = expenseCheck.checkDocument(listOfExpense, pathDirectoryInvoice);
        expenseCheck.writeFileListOfExpense(listOfExpense, checkDocument);

        Ledger ledger = new Ledger(codeCondominium);
        InfoGrandLivre infoGrandLivre = ledger.getInfoGrandLivre(pathFileLeger);
        if (infoGrandLivre.stopDate().getYear() != accountingYear) {
            LOGGER.error("L'année comptable n'est pas cohérente (dans le grand livre : {}, passée en paramètre : {})", infoGrandLivre.stopDate(), accountingYear);
        }

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts(pathFileLeger, codeCondominium);
        account.writeFilesAccounts(accounts, infoGrandLivre, PATH);

        List<Line> lineBankInGrandLivre = ledger.getInfoBankGrandLivre(infoGrandLivre, accounts, pathFileLeger,
                pathDirectoryInvoice, accountsbank);

        // Récupération des relevés bancaire
        Bank bank = new Bank();
        List<BankLine> bankLines = bank.getBankLines(accounts, pathDirectoryBank, accountsbank,
                accountingYear);

        String nameFile = PATH + "Etat de rapprochement.xlsx";
        WriteFile writeFile = new WriteFile(PATH);
        writeFile.writeFileExcelEtatRaprochement(lineBankInGrandLivre, nameFile, bankLines);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}