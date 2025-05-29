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

    private static final int NAME_JOURNAL_SIZE = 2;
    private static final String REGEX_NUMBER = "^-?[0-9]+$";
    private static final String VIRT = "Virt";
    private static final String TOTAL_IMMEUBLE = "Total immeuble";
    private static final String GRAND_LIVRE = "GrandLivre";
    private static final String DIRECTORY_NAME_COPROPRIETAIRE = "Copropriétaire";
    public static final String TOTAL_COMPTE = "Total compte";
    public static final String SOLDE = "Solde";
    private static String codeCondominium;
    private static String postalCode;
    private final OutilInfo outilInfo = new OutilInfo();
    private final Constant constant = new Constant();

    public Ledger(String codeCondominium) {
        Ledger.codeCondominium = codeCondominium;
    }

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
            Ledger.postalCode = postalCode;
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
        if (firstWord.equals(postalCode)) {
            return false;
        }
        if (firstWord.equals(codeCondominium)) {
            return false;
        }
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

        // Calcule de nombre de montants sur la ligne grâce au signe €
        long numberOfAmounts = outilInfo.getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);

        int indexOfWords = 0;

        String document = outilInfo.getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }

        indexOfWords = outilInfo.getIndexOfNextWords(words, indexOfWords);

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
            account = outilInfo.getAccount(accounts, numberAccounts, 0);
            if (account == null) {
                String[] numAccount = {words[indexOfWords] + words[indexOfWords + 1]};
                account = outilInfo.getAccount(accounts, numAccount, 0);
            }
            if (account == null) {
                constant.logError(line, words[indexOfWords]);
                return null;
            }
            indexOfWords++;
        } else {
            account = outilInfo.getAccount(accounts, words, indexOfWords);
            if (account == null) {
                String[] numAccount = {words[indexOfWords] + words[indexOfWords + 1]};
                account = outilInfo.getAccount(accounts, numAccount, 0);
                indexOfWords++;
            }
            if (account == null) {
                constant.logError(ERREUR_LORS_DE_LA_RECHERCHE_DU_COMPTE_SUR_LA_LIGNE, line, words[indexOfWords]);
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
        if (!line.contains(REPORT_DE) && words[indexOfWords].replace("-", "").matches(REGEX_NUMBER)) {
            accountCounterpart = outilInfo.getAccount(accounts, words, indexOfWords);
            indexOfWords++;
        }

        if (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }

        // Extraction du numéro de chéque
        String checkNumber = "";
        if (VIRT.equals(words[indexOfWords])) {
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
            indexOfWords = getIndexOfWords(words, indexOfWords, label);
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            indexOfWords = getIndexOfWords(words, indexOfWords, debit);
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            indexOfWords = getIndexOfWords(words, indexOfWords, credit);
            credit.append(" ").append(words[indexOfWords]);
        } else if (numberOfAmounts == 2 && words[indexOfWords].contains("Report")) {
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

        debit = new StringBuilder(debit.toString().replace(" ", "").replace(Character.toString(EURO), "").trim());
        credit = new StringBuilder(credit.toString().replace(" ", "").replace(Character.toString(EURO), "").trim());

        return new Line(document, date, account, journal, accountCounterpart, checkNumber,
                label.toString().trim(), debit.toString().trim(), credit.toString().trim());
    }

    private int getIndexOfWords(String[] words, int indexOfWords, StringBuilder label) {
        while (outilInfo.isAmount(words[indexOfWords])) {
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
        }
        return indexOfWords;
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
        line = line.replace("| ", " ");
        // Correction des espaces avant le signe €
        line = outilInfo.fixedSpacesBeforeEuroSign(line);

        // Calcule de nombre de montants sur la ligne grâce au signe €
        long numberOfAmounts = outilInfo.getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);
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
            indexOfWords = getIndexOfWords(words, indexOfWords, debit);
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            indexOfWords = getIndexOfWords(words, indexOfWords, credit);
            credit.append(" ").append(words[indexOfWords]);
        } else if (numberOfAmounts >= 2) {
            indexOfWords = getIndexOfWords(words, indexOfWords, debit);
            debit.append(" ").append(words[indexOfWords]);
            credit = debit;
        } else {
            constant.logError(ERREUR_IL_MANQUE_DES_MONTANTS_SUR_LA_LIGNE_DE_TOTAL, line);
        }

        // Extraction du numéro de compte
        String[] labels = outilInfo.splittingLineIntoWordTable(label.toString().replace(Ledger.TOTAL_COMPTE, "").trim());
        TypeAccount account = null;
        for (int indexOfLabel = 0; indexOfLabel < labels.length; indexOfLabel++) {
            account = outilInfo.getAccount(accounts, labels, indexOfLabel);
            if (account != null) {
                label = new StringBuilder(label.toString().replace(account.account(), "").replace("  ", " "));
                break;
            }
        }
        if (account == null) {
            StringBuilder numAccount = new StringBuilder();
            StringBuilder accountInLabel = new StringBuilder();
            for (String word : labels) {
                if (word.contains(SOLDE)) {
                    break;
                }
                numAccount.append(word);
                accountInLabel.append(" ").append(word);
            }
            if (!numAccount.isEmpty()) {
                String[] numAccountWord = new String[1];
                numAccountWord[0] = numAccount.toString();
                account = outilInfo.getAccount(accounts, numAccountWord, 0);
                if (account != null) {
                    label = new StringBuilder(label.toString().replace(accountInLabel.toString(), "").replace("  ", " "));
                }
            }
        }
        if (account == null) {
            constant.logError(ERREUR_LORS_DE_LA_RECHERCHE_DU_COMPTE_SUR_LE_TOTAL_DU_COMPTE_DE_LA_LIGNE, label.toString(), line);
            return null;
        }

        debit = new StringBuilder(debit.toString().replace(" ", "").replace(Character.toString(EURO), "").trim());
        credit = new StringBuilder(credit.toString().replace(" ", "").replace(Character.toString(EURO), "").trim());

        return new TotalAccount(label.toString().trim(), account, debit.toString(), credit.toString().replace("—", "").trim());
    }

    public boolean isTotalAccount(String line) {
        return line.startsWith(TOTAL_COMPTE);
    }

    private String postalCode(String line) {
        return line.split(" ")[0];
    }

    public boolean isTotalBuilding(String line) {
        return line.startsWith(TOTAL_IMMEUBLE);
    }

    public TotalBuilding totalBuilding(String line) {
        // Correction des espaces avant le signe €
        line = outilInfo.fixedSpacesBeforeEuroSign(line).replace(")", " ) ");

        // Calcule de nombre de montants sur la ligne grâce au signe €
        long numberOfAmounts = outilInfo.getNumberOfAmountsOn(line);

        // Découpage de la ligne en un tableau de mot
        String[] words = outilInfo.splittingLineIntoWordTable(line);
        int indexOfWords = 0;
        StringBuilder label = new StringBuilder();
        StringBuilder debit = new StringBuilder();
        StringBuilder credit = new StringBuilder();
        if (numberOfAmounts == 3) {
            indexOfWords = getIndexOfWords(words, indexOfWords, label);
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            label.append(" ").append(words[indexOfWords]);
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords = getIndexOfWords(words, indexOfWords, debit);
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            indexOfWords = getIndexOfWords(words, indexOfWords, credit);
            credit.append(" ").append(words[indexOfWords]);
        }
        return new TotalBuilding(label.toString().trim().replace(" )", ")"), debit.toString().trim().replace(" ", "").replace(Character.toString(EURO), ""), credit.toString().trim().replace(" ", "").replace(Character.toString(EURO), ""));
    }

    public List<Line> getInfoBankGrandLivre(InfoGrandLivre infoGrandLivre, Map<String, TypeAccount> accounts,
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

        WriteFile writeFile = new WriteFile();
        String pathFile = "." + File.separator + DIRECTORY_NAME_RESULTAT + File.separator;
        String exitFile = pathFile + GRAND_LIVRE + CSV;
        writeFile.writeFileCSVGrandLivre(grandLivres, exitFile);
        String nameFile = infoGrandLivre.printDate() + " " + GRAND_LIVRE + " " + infoGrandLivre.syndicName()
                + " au " + infoGrandLivre.stopDate() + XLSX;
        String path = pathFile + nameFile;
        writeFile.writeFileExcelGrandLivre(grandLivres, path, journals, pathDirectoryInvoice);

        writeFile.writeFilesExcelCoOwner(grandLivres, pathFile + DIRECTORY_NAME_COPROPRIETAIRE + File.separator, accounts, pathDirectoryInvoice);
        return lineBankInGrandLivre;
    }
}
