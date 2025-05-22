package life.light;

import life.light.extract.info.Ledger;
import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.temporal.ChronoUnit;
import java.util.*;

import static life.light.extract.info.Account.getAccounts;
import static life.light.extract.info.Bank.getBankLines;
import static life.light.write.WriteFile.*;
import static life.light.extract.info.Ledger.getInfoGrandLivre;
import static life.light.extract.info.Ledger.getNumberOfLineInFile;

public class Main {

    private static final Logger LOGGER = LogManager.getLogger();
    // TODO : à mettre dans les paramètres de lancement du programme
    public static String BANK_1_ACCOUNT = "51220";
    public static String BANK_2_ACCOUNT = "51221";
    public static String PATH_DIRECTORY_INVOICE = "";
    public static String PATH_DIRECTORY_BANK = "";
    public static String PATH_DIRECTORY_LEDGER = "";
    public static String codeCondominium = "";

    public static void main(String[] args) {
        LocalDateTime debut = LocalDateTime.now();
        LOGGER.info("Début à {}:{}:{}", debut.getHour(), debut.getMinute(), debut.getSecond());

        if (args.length == 1) {
            codeCondominium = args[0];
        }
        InfoGrandLivre infoGrandLivre = getInfoGrandLivre(PATH_DIRECTORY_LEDGER);
        LOGGER.info("Le nom du syndic est : {}", infoGrandLivre.syndicName());
        LOGGER.info("La date d'édition est le {}", infoGrandLivre.printDate());
        LOGGER.info("La date d'arrêt des comptes est le {}", infoGrandLivre.stopDate());
        LOGGER.info("Le code postal du syndic est {}", infoGrandLivre.postalCode());

        Map<String, TypeAccount> accounts = getAccounts(PATH_DIRECTORY_LEDGER, infoGrandLivre.postalCode(), codeCondominium);

        writeFileCSVAccounts(accounts, "." + File.separator + "temp" + File.separator + "Plan comptable de " + infoGrandLivre.syndicName() + ".csv");
        writeFileExcelAccounts(accounts, "." + File.separator + "temp" + File.separator + "Plan comptable de " + infoGrandLivre.syndicName() + ".xlsx");

        int numberOfLineInFile = getNumberOfLineInFile(PATH_DIRECTORY_LEDGER);

        // Récupération des relevés bancaire
        List<String> pathsDirectoryBank = new ArrayList<>();
        pathsDirectoryBank.add(PATH_DIRECTORY_BANK + File.separator + BANK_2_ACCOUNT);
        pathsDirectoryBank.add(PATH_DIRECTORY_BANK + File.separator + BANK_1_ACCOUNT);
        List<BankLine> bankLines = getBankLines(accounts, pathsDirectoryBank, infoGrandLivre.stopDate().getYear());

        // Géneration du grand livre
        Object[] grandLivres = new Object[numberOfLineInFile];
        TreeSet<String> journals = new TreeSet<>();
        List<Line> lineBankInGrandLivre = new ArrayList<>();
        int indexInGrandLivres = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(PATH_DIRECTORY_LEDGER))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.isEmpty()) {
                    continue;
                }
                if (Ledger.isLigne(line)) {
                    Line lineOfGrandLivre = Ledger.line(line, accounts);
                    if (lineOfGrandLivre == null) {
                        continue;
                    }
                    grandLivres[indexInGrandLivres++] = lineOfGrandLivre;
                    if (!lineOfGrandLivre.journal().isEmpty()) {
                        journals.add(lineOfGrandLivre.journal());
                    }
                    String accountNumber = lineOfGrandLivre.account().account();
                    if (accountNumber.equals(BANK_2_ACCOUNT) || 
                        (accountNumber.equals(BANK_1_ACCOUNT) && !lineOfGrandLivre.label().contains("Report de "))) {
                        lineBankInGrandLivre.add(lineOfGrandLivre);
                    }
                } else if (Ledger.isTotalAccount(line)) {
                    TotalAccount totalAccount = Ledger.totalAccount(line, accounts);
                    if (totalAccount != null) {
                        grandLivres[indexInGrandLivres++] = totalAccount;
                    }
                } else if (Ledger.isTotalBuilding(line)) {
                    TotalBuilding totalBuilding = Ledger.totalBuilding(line);
                    grandLivres[indexInGrandLivres++] = totalBuilding;
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }

        writeFileCSVGrandLivre(grandLivres);
        String nameFile = infoGrandLivre.printDate().substring(6) + "-" + infoGrandLivre.printDate().substring(3, 5) + "-" + infoGrandLivre.printDate().substring(0, 2)
                + " Grand livre " + infoGrandLivre.syndicName().substring(0, infoGrandLivre.syndicName().length() - 1).trim()
                + " au " + infoGrandLivre.stopDate()
                + ".xlsx";
        writeFileExcelGrandLivre(grandLivres, nameFile, journals);

        nameFile = "." + File.separator + "temp" + File.separator + "Etat de rapprochement.xlsx";
        writeFileExcelEtatRaprochement(lineBankInGrandLivre, nameFile, bankLines);

        LocalDateTime fin = LocalDateTime.now();
        LOGGER.info("La durée du traitement est de {} secondes", ChronoUnit.SECONDS.between(debut, fin));
    }
}