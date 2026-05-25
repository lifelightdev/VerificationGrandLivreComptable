package life.light;

import life.light.extract.info.Account;
import life.light.extract.info.Bank;
import life.light.extract.info.Ledger;
import life.light.type.BankLine;
import life.light.type.Line;
import life.light.type.TypeAccount;
import life.light.write.WriteFile;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.List;
import java.util.Map;

import static life.light.Constant.PATH;

public class Main {

    private static final Logger LOGGER = LogManager.getLogger();
    static List<String> accountsbank = List.of("51220", "51221");
    static String pathDirectoryInvoice = "";
    static String pathDirectoryBank = "";
    static String pathFileLeger = "GRAND_LIVRE.csv";
    static String pathFileListOfExpenses = "";
    static String codeCondominium = "";
    static int accountingYear = 2025;

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

        Ledger ledger = new Ledger();

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts();
        account.writeFilesAccounts(accounts, PATH);

        List<Line> lineBankInGrandLivre = ledger.getInfoBankGrandLivre(accounts, pathFileLeger, pathDirectoryInvoice, accountsbank);

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