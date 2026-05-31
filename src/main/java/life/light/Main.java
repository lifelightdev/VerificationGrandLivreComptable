package life.light;

import life.light.check.OutilChek;
import life.light.extract.info.Account;
import life.light.extract.info.Bank;
import life.light.extract.info.Ledger;
import life.light.type.*;
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
    static List<String> accountsBank = List.of("51221");
    static String pathDirectoryInvoice = "";
    static String pathDirectoryBank = "";
    static String pathFileLeger = "GRAND_LIVRE.csv";
    static String pathFileListOfExpenses = "";
    static String codeCondominium = "";
    static int accountingYear = 2025;

    static void main() {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());

        Ledger ledger = new Ledger();

        Account account = new Account();
        Map<String, TypeAccount> accounts = account.getAccounts();
        account.writeFilesAccounts(accounts, PATH);

        List<LineLedger> lineLedgerBankInGrandLivre = ledger.getInfoBankGrandLivre(accounts, pathFileLeger, pathDirectoryInvoice, accountsBank);

        // Récupération des relevés bancaire
        Bank bank = new Bank();
        List<BankLine> bankLines = bank.getBankLines(accounts, pathDirectoryBank, accountsBank,
                accountingYear);

        StateOfReconciliation stateOfReconciliation = OutilChek.getStateOfConvergence(lineLedgerBankInGrandLivre, bankLines);

        String nameFile = PATH + "État de rapprochement V1.xlsx";
        WriteFile writeFile = new WriteFile(PATH);
        writeFile.writeFileExcelEtatRaprochement(lineLedgerBankInGrandLivre, nameFile, bankLines);
        nameFile = PATH + "État de rapprochement V2.xlsx";
        writeFile.writeFileExcelStateOfReconciliation(stateOfReconciliation, nameFile);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}