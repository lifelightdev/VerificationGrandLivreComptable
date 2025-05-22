package life.light.extract.info;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import static life.light.Main.*;
import static life.light.extract.info.OutilInfo.*;
import static life.light.write.WriteFile.writeFileCSVGrandLivre;
import static life.light.write.WriteFile.writeFileExcelGrandLivre;

public class Ledger {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final int NAME_JOURNAL_SIZE = 2;
    public static final String REGEX_NUMBER = "^-?[0-9]+$";

    public InfoGrandLivre getInfoGrandLivre(String pathLedger) {
        String syndicName = "";
        String printDate = "";
        LocalDate stopDate = null;
        String postalCode = "";
        // Récupération des informations pour la génération du nom de fichier
        try (BufferedReader reader = new BufferedReader(new FileReader(pathLedger))) {
            String line = reader.readLine();
            syndicName = syndicName(line);
            line = reader.readLine();
            printDate = printDate(line);
            reader.readLine();
            line = reader.readLine();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            stopDate = LocalDate.parse(stopDate(line), formatter);
            postalCode = postalCode(line);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Le nom du syndic est : {}", syndicName.trim());
        LOGGER.info("La date d'édition est le {}", printDate);
        LOGGER.info("La date d'arrêt des comptes est le {}", stopDate);
        LOGGER.info("Le code postal du syndic est {}", postalCode);
        return new InfoGrandLivre(syndicName.trim(), printDate, stopDate, postalCode);
    }

    public int getNumberOfLineInFile(String pathLedger) {
        int numberOfLineInFile = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(pathLedger))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (!line.trim().isEmpty()) {
                    numberOfLineInFile++;
                }
            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        LOGGER.info("Il y a {} lignes dans le grand livre", numberOfLineInFile);
        return numberOfLineInFile;
    }

    protected String syndicName(String ligne) {
        return ligne.replace("|", "");
    }

    String printDate(String line) {
        return findDateIn(line);
    }

    String stopDate(String line) {
        return findDateIn(line);
    }

    TypeAccount account(String line) {
        line = getLine(line);
        String[] words = splittingLineIntoWordTable(line);
        String account = words[0];
        StringBuilder label = new StringBuilder();
        for (String word : words) {
            if (!account.equals(word)) {
                label.append(" ").append(word);
            }
        }
        if (label.toString().contains("*")) {
            int index = label.toString().indexOf("*");
            label = new StringBuilder(label.substring(0, index));
        }
        return new TypeAccount(account, label.toString().trim());
    }

    private String getLine(String line) {
        line = line.replace(" _ ", " ");
        line = line.replace("_ ", " ");
        line = line.replace(" | ", " ");
        line = line.replace(" — ", " ");
        line = line.replace(" … ", " ");
        line = line.replace("°", " ");
        line = line.replace("#", " ");
        line = line.replace("‘", " ");
        line = line.replace("=—", " ");
        line = line.replace("4A01", "401");
        return line;
    }

    boolean isAcccount(String line, String postalCode, String id) {
        line = getLine(line);
        line = line.replace("  ", " ");
        if (line.contains(EURO)) {
            return false;
        }
        if ((line.contains("Page")) && (!findDateIn(line).isEmpty())) {
            return false;
        }
        if (line.trim().isEmpty()) {
            return false;
        }
        String[] words = splittingLineIntoWordTable(line);
        String firstWord = words[0];
        if (firstWord.equals(postalCode)) {
            return false;
        }
        if (firstWord.equals(id)) {
            return false;
        }
        if (firstWord.contains("-")) {
            firstWord = firstWord.replace("-", "");
        }
        if (firstWord.matches(REGEX_NUMBER) || "461VC".equals(firstWord)) {
            return findDateIn(words[1]).isEmpty();
        }
        return false;
    }

    public Line line(String line, Map<String, TypeAccount> accounts) {
        // Supprime les caractères parasites
        line = removesStrayCharactersInLine(line);

        // Correction des espaces avant le signe €
        line = fixedSpacesBeforeEuroSign(line);

        // Calcule de nombre de montants sur la ligne grâce au signe €
        int numberOfAmounts = getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = splittingLineIntoWordTable(line);

        int indexOfWords = 0;

        String document = getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }

        indexOfWords = getIndexOfWords(words, indexOfWords);

        // Extraction de la date de l'opération
        String date = words[indexOfWords];
        indexOfWords++;

        TypeAccount account;
        String journal = "";
        // Extraction du numéro de compte
        if (words[indexOfWords].length() > 10) {
            String numberAccount = words[indexOfWords].substring(0, 10);
            journal = words[indexOfWords].substring(10);
            String[] numberAccounts = {numberAccount};
            account = getAccount(accounts, numberAccounts, 0);
            if (account == null) {
                String[] numAccount = {words[indexOfWords] + words[indexOfWords + 1]};
                account = getAccount(accounts, numAccount, 0);
            }
            if (account == null) {
                LOGGER.error("Erreur lors de la recherche du compte {} sur la ligne {}", words[indexOfWords], line);
                return null;
            }
            indexOfWords++;
        } else {
            account = getAccount(accounts, words, indexOfWords);
            if (account == null) {
                String[] numAccount = {words[indexOfWords] + words[indexOfWords + 1]};
                account = getAccount(accounts, numAccount, 0);
                indexOfWords++;
            }
            if (account == null) {
                LOGGER.error("Erreur lors de la recherche du compte {} sur la ligne {}", words[indexOfWords], line);
                return null;
            }
            indexOfWords++;
            if (words[indexOfWords].trim().isEmpty()) {
                indexOfWords++;
            }
            // Extraction du code du journal
            if (words[indexOfWords].length() == NAME_JOURNAL_SIZE) {
                journal = words[indexOfWords];
                indexOfWords++;
            }
        }

        if (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }

        // Extraction du compte de contrepartie
        TypeAccount accountCounterpart = null;
        if (!line.contains("Report de") && words[indexOfWords].replace("-", "").matches(REGEX_NUMBER)) {
            accountCounterpart = getAccount(accounts, words, indexOfWords);
            indexOfWords++;
        }

        if (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }

        // Extraction du numéro de chéque
        String checkNumber = "";
        if ("Virt".equals(words[indexOfWords])) {
            checkNumber = words[indexOfWords];
            indexOfWords++;
        }

        if (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }

        // Extraction du libellé et des montants
        StringBuilder label = new StringBuilder();
        StringBuilder debit = new StringBuilder();
        StringBuilder credit = new StringBuilder();
        if (numberOfAmounts == 3) {
            // S'il y a trois montants, c'est un report donc
            // le premier montant est dans le libellé
            // le second montant est le débit
            // le troisème montant est le crédit
            while (!words[indexOfWords].endsWith(EURO)) {
                label.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].endsWith(EURO)) {
                debit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].endsWith(EURO)) {
                credit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            credit.append(" ").append(words[indexOfWords]);
        } else if (numberOfAmounts == 2 && words[indexOfWords].contains("Report")) {
            // S'il y a deux montants et que c'est un report
            int indexOfWordsStartEnd = words.length - 1;
            StringBuilder amount = new StringBuilder(words[indexOfWordsStartEnd]);
            indexOfWordsStartEnd--;
            while (words[indexOfWordsStartEnd].replace(".", "").matches(REGEX_NUMBER)) {
                amount.insert(0, words[indexOfWordsStartEnd] + " ");
                indexOfWordsStartEnd--;
            }
            for (int i = indexOfWords; i <= indexOfWordsStartEnd; i++) {
                label.append(" ").append(words[i]);
            }

            if (label.toString().endsWith(" ")) {
                credit = amount;
            } else {
                debit = amount;
            }
        } else if (numberOfAmounts == 1) {
            int indexOfWordsStartEnd = words.length - 1;
            if (words[indexOfWordsStartEnd].trim().isEmpty()) {
                indexOfWordsStartEnd--;
            }
            StringBuilder amount = new StringBuilder(words[indexOfWordsStartEnd]);
            indexOfWordsStartEnd--;
            while (words[indexOfWordsStartEnd].replace(".", "").matches(REGEX_NUMBER)) {
                String word = words[indexOfWordsStartEnd];
                if (!"2024".equals(word)) {
                    if ((word.length() != 9)) {
                        amount.insert(0, word + " ");
                        indexOfWordsStartEnd--;
                    } else {
                        break;
                    }
                } else {
                    break;
                }
            }
            for (int i = indexOfWords; i <= indexOfWordsStartEnd; i++) {
                label.append(" ").append(words[i]);
            }
            if (!words[words.length - 1].trim().isEmpty()) {
                credit = amount;
            } else {
                debit = amount;
            }
        } else if (numberOfAmounts == 0) {
            while (indexOfWords < words.length) {
                label.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
        }

        if (!debit.toString().contains(".") && !debit.toString().isEmpty()) {
            debit = new StringBuilder(debit.substring(0, debit.toString().length() - 3) + "." + debit.substring(debit.toString().length() - 3));
        }

        debit = new StringBuilder(debit.toString().replace(" ", "").replace("€", "").trim());
        credit = new StringBuilder(credit.toString().replace(" ", "").replace("€", "").trim());

        return new Line(document, date, account, journal, accountCounterpart, checkNumber,
                label.toString().trim(), debit.toString().trim(), credit.toString().trim());
    }

    public boolean isLigne(String line) {
        line = removesStrayCharactersInLine(line);
        // Découpage de la ligne en un tableau de mot
        String[] words = splittingLineIntoWordTable(line);
        int indexOfWords = 0;
        // Vérification que la ligne commence par une pièce
        String document = getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }
        while (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }
        /* Vérification que la ligne commence par une date ou qu'il y a une date après la pièce */
        return !findDateIn(words[indexOfWords]).isEmpty();
    }

    public TotalAccount totalAccount(String line, Map<String, TypeAccount> accounts) {
        line = line.replace("| ", " ");
        // Correction des espaces avant le signe €
        line = fixedSpacesBeforeEuroSign(line);

        // Calcule de nombre de montants sur la ligne grâce au signe €
        int numberOfAmounts = getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = splittingLineIntoWordTable(line);
        int indexOfWords = 0;
        StringBuilder label = new StringBuilder();
        while (!words[indexOfWords].endsWith(")")) {
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
        }
        label.append(" ").append(words[indexOfWords]);
        indexOfWords++;

        StringBuilder debit = new StringBuilder();
        StringBuilder credit = new StringBuilder();
        if (numberOfAmounts == 3) {
            while (!words[indexOfWords].endsWith(EURO)) {
                debit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].endsWith(EURO)) {
                credit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            credit.append(" ").append(words[indexOfWords]);
        } else if (numberOfAmounts >= 2) {
            while (!words[indexOfWords].endsWith(EURO)) {
                debit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            debit.append(" ").append(words[indexOfWords]);
            credit = debit;
        } else {
            LOGGER.error("Erreur, il manque des montants sur la ligne de total {}", line);
        }

        // Extraction du numéro de compte
        String[] labels = splittingLineIntoWordTable(label.toString().replace("Total compte ", "").trim());
        TypeAccount account = null;
        for (int indexOfLabel = 0; indexOfLabel < labels.length; indexOfLabel++) {
            account = getAccount(accounts, labels, indexOfLabel);
            if (account != null) {
                label = new StringBuilder(label.toString().replace(account.account(), "").replace("  ", " "));
                break;
            }
        }
        if (account == null) {
            StringBuilder numAccount = new StringBuilder();
            StringBuilder accountInLabel = new StringBuilder();
            for (String word : labels) {
                if (word.contains("Solde")) {
                    break;
                }
                numAccount.append(word);
                accountInLabel.append(" ").append(word);
            }
            if (!numAccount.isEmpty()) {
                String[] numAccountWord = new String[1];
                numAccountWord[0] = numAccount.toString();
                account = getAccount(accounts, numAccountWord, 0);
                if (account != null) {
                    label = new StringBuilder(label.toString().replace(accountInLabel.toString(), "").replace("  ", " "));
                }
            }
        }
        if (account == null) {
            LOGGER.error("Erreur lors de la recherche du compte sur le total du compte {} de la ligne {}", label, line);
            return null;
        }

        debit = new StringBuilder(debit.toString().replace(" ", "").replace("€", "").trim());
        credit = new StringBuilder(credit.toString().replace(" ", "").replace("€", "").trim());

        return new TotalAccount(label.toString().trim(), account, debit.toString(), credit.toString().replace("—", "").trim());
    }

    public boolean isTotalAccount(String line) {
        return line.startsWith("Total compte ");
    }

    String postalCode(String line) {
        return line.split(" ")[0];
    }

    public boolean isTotalBuilding(String line) {
        return line.startsWith("Total immeuble");
    }

    public TotalBuilding totalBuilding(String line) {
        // Correction des espaces avant le signe €
        line = fixedSpacesBeforeEuroSign(line).replace(")", " ) ");

        // Calcule de nombre de montants sur la ligne grâce au signe €
        int numberOfAmounts = getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = splittingLineIntoWordTable(line);
        int indexOfWords = 0;
        StringBuilder label = new StringBuilder();
        StringBuilder debit = new StringBuilder();
        StringBuilder credit = new StringBuilder();
        if (numberOfAmounts == 3) {
            while (!words[indexOfWords].endsWith(EURO)) {
                label.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            label.append(" ").append(words[indexOfWords]);
            debit.append(" ").append(words[indexOfWords]);
            while (!words[indexOfWords].endsWith(EURO)) {
                debit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].endsWith(EURO)) {
                credit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            credit.append(" ").append(words[indexOfWords]);
        }
        return new TotalBuilding(label.toString().trim().replace(" )", ")"), debit.toString().trim().replace(" ", "").replace(EURO, ""), credit.toString().trim().replace(" ", "").replace(EURO, ""));
    }

    public List<Line> getInfoBankGrandLivre(InfoGrandLivre infoGrandLivre, Map<String, TypeAccount> accounts) {
        // Géneration du grand livre
        Object[] grandLivres = new Object[getNumberOfLineInFile(PATH_DIRECTORY_LEDGER)];
        TreeSet<String> journals = new TreeSet<>();
        List<Line> lineBankInGrandLivre = new ArrayList<>();
        int indexInGrandLivres = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(PATH_DIRECTORY_LEDGER))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.isEmpty()) {
                    continue;
                }
                if (isLigne(line)) {
                    Line lineOfGrandLivre = line(line, accounts);
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
                } else if (isTotalAccount(line)) {
                    TotalAccount totalAccount = totalAccount(line, accounts);
                    if (totalAccount != null) {
                        grandLivres[indexInGrandLivres++] = totalAccount;
                    }
                } else if (isTotalBuilding(line)) {
                    TotalBuilding totalBuilding = totalBuilding(line);
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
        return lineBankInGrandLivre;
    }
}
