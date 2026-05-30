package life.light.extract.info;

import life.light.Constant;
import life.light.type.*;
import life.light.write.WriteFile;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import static life.light.Constant.*;

public class Ledger {

    private static final String REGEX_NUMBER = "^-?[0-9]+$";
    private static final String GRAND_LIVRE = "GrandLivre";
    private static final String DIRECTORY_NAME_COPROPRIETAIRE = "Copropriétaire";
    public static final String TOTAL_COMPTE = "Total compte";
    private final OutilInfo outilInfo = new OutilInfo();
    private final Constant constant = new Constant();

    public InfoGrandLivre getInfoGrandLivre(String pathLedger) {
        try (BufferedReader reader = new BufferedReader(new FileReader(pathLedger))) {
            String line = reader.readLine();
            String syndicName = line.replace("|", "").trim();
            line = reader.readLine();
            LocalDate printDate = date(line);
            reader.readLine();
            line = reader.readLine();
            LocalDate stopDate = date(line);
            String postalCode = postalCode(line);
            return new InfoGrandLivre(syndicName, printDate, stopDate, postalCode);
        } catch (IOException e) {
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
        }
        return null;
    }

    private LocalDate date(String line) {
        String date = outilInfo.findDateIn(line);
        if (date.isEmpty()) {
            return null;
        }
        return LocalDate.parse(outilInfo.findDateIn(line), DATE_FORMATTER);
    }

    TypeAccount account(String line) {
        line = getLine(line);
        String[] words = outilInfo.splittingLineIntoWordTable(line);
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

    boolean isAccount(String line) {
        line = getLine(line);
        line = line.replace("  ", " ");
        if (line.contains(Character.toString(EURO))) {
            return false;
        }
        if ((line.contains("Page")) && (!outilInfo.findDateIn(line).isEmpty())) {
            return false;
        }
        if (line.trim().isEmpty()) {
            return false;
        }
        String[] words = outilInfo.splittingLineIntoWordTable(line);
        String firstWord = words[0];
        if (firstWord.contains("-")) {
            firstWord = firstWord.replace("-", "");
        }
        if (firstWord.matches(REGEX_NUMBER) || "461VC".equals(firstWord)) {
            return outilInfo.findDateIn(words[1]).isEmpty();
        }
        return false;
    }

    public Line line(String line, Map<String, TypeAccount> accounts) {
        // Supprime les caractères parasites
        line = outilInfo.removesStrayCharactersInLine(line);

        // Correction des espaces avant le signe €
        line = outilInfo.fixedSpacesBeforeEuroSign(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);

        String document = words[0];
        // Extraction de la date de l'opération
        String date = words[1];

        TypeAccount account;
        // Extraction du numéro de compte
        String numberAccount = words[2].trim();
        account = outilInfo.getAccount(accounts, numberAccount);
        if (account == null) {
            account = outilInfo.getAccount(accounts, numberAccount);
        }
        // Extraction du journal
        String journal = words[3];

        // Extraction du compte de contrepartie
        TypeAccount accountCounterpart = null;
        if (!line.contains(REPORT_DE) && words[4].replace("-", "").matches(REGEX_NUMBER)) {
            accountCounterpart = outilInfo.getAccount(accounts, words[4]);
            if (accountCounterpart == null) {
                LOGGER.info("accountCounterpart is null à la ligne : {}", line);
            }
        }

        // Extraction du numéro de chéque
        String checkNumber = words[5];

        // Extraction du libellé et des montants
        String label;
        String debit = "";
        String credit = "";
        label = words[6];
        if (words.length > 7) {
            debit = words[7];
        }
        if (words.length > 8) {
            credit = words[8];
        }

        if (label.startsWith("0 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("1 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("3 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("4 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("5 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("6 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("7 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("8 ")) {
            label = label.substring(1).trim();
        }
        if (label.startsWith("10 ")) {
            label = label.substring(2).trim();
        }
        if (label.startsWith("00 ")) {
            label = label.substring(2).trim();
        }
        if (label.startsWith("20 ")) {
            label = label.substring(2).trim();
        }
        if (label.startsWith("41 ")) {
            label = label.substring(2).trim();
        }

        return new Line(document.trim(), date.trim(), account, journal.trim(), accountCounterpart, checkNumber.trim(),
                label.trim(), debit.trim(), credit.trim());
    }

    public boolean isLigne(String line) {
        line = outilInfo.removesStrayCharactersInLine(line);
        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);
        int indexOfWords = 0;
        // Vérification que la ligne commence par une pièce
        String document = outilInfo.getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }
        while (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }
        /* Vérification que la ligne commence par une date ou qu'il y a une date après la pièce */
        return !outilInfo.findDateIn(words[indexOfWords]).isEmpty();
    }

    public TotalAccount totalAccount(String line, Map<String, TypeAccount> accounts) {
        // Correction des espaces avant le signe €
        line = outilInfo.fixedSpacesBeforeEuroSign(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);

        String label = words[6];
        String debit = words[7];
        String credit = words[8];

        // Extraction du numéro de compte
        String[] labels = label.split(" ");
        TypeAccount account = null;
        for (String s : labels) {
            account = outilInfo.getAccount(accounts, s);
            if (account != null) {
                label = (label.replace(account.account(), "").replace("  ", " "));
                break;
            }
        }

        if (account == null) {
            constant.logError(ERREUR_LORS_DE_LA_RECHERCHE_DU_COMPTE_SUR_LE_TOTAL_DU_COMPTE_DE_LA_LIGNE, label, line);
            return null;
        }
        return new TotalAccount(label.trim(), account, debit.trim(), credit.trim());
    }

    public boolean isTotalAccount(String line) {
        return line.contains(TOTAL_COMPTE);
    }

    private String postalCode(String line) {
        return line.split(" ")[0];
    }

    public boolean isTotalBuilding(String line) {
        return line.contains(TOTAL_IMMEUBLE);
    }

    public TotalBuilding totalBuilding(String line) {
        // Correction des espaces avant le signe €
        line = outilInfo.fixedSpacesBeforeEuroSign(line).replace(")", " ) ");

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);
        String label = words[6];
        String debit = words[7];
        String credit = words[8];

        return new TotalBuilding(label.trim().replace(" )", ")"), debit.trim().replace(" ", "").replace(Character.toString(EURO), ""), credit.trim().replace(" ", "").replace(Character.toString(EURO), ""));
    }

    public List<Line> getInfoBankGrandLivre(Map<String, TypeAccount> accounts,
                                            String pathDirectoryLeger, String pathDirectoryInvoice,
                                            List<String> accountsbank) {
        // Géneration du grand livre
        Object[] grandLivres = new Object[outilInfo.getNumberOfLineInFile(pathDirectoryLeger)];
        TreeSet<String> journals = new TreeSet<>();
        List<Line> lineBankInGrandLivre = new ArrayList<>();
        int indexInGrandLivres = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(pathDirectoryLeger))) {
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
                    if (accountsbank.contains(accountNumber) && !lineOfGrandLivre.label().contains(REPORT_DE)) {
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
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
        }

        WriteFile writeFile = new WriteFile(PATH);
        String nameFile = GRAND_LIVRE + XLSX;
        String path = PATH + nameFile;
        writeFile.writeFileExcelGrandLivre(grandLivres, path, journals, pathDirectoryInvoice);

        writeFile.writeFilesExcelCoOwner(grandLivres, PATH + DIRECTORY_NAME_COPROPRIETAIRE + File.separator, accounts, pathDirectoryInvoice);
        return lineBankInGrandLivre;
    }
}
