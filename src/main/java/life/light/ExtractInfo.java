package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.Map;

public class ExtractInfo {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final int NAME_JOURNAL_SIZE = 2;
    public static final String EURO = "€";
    public static final String REGEX_IS_DATE = "^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(19|20)\\d{2}$";
    public static final String REGEX_NUMBER = "^-?[0-9]+$";
    public static final String REGEX_PHONE_NUMBER = "^0[1-9]([-. ]?\\d{2}){4}$";

    public static String syndicName(String ligne) {
        return ligne.replace("|", "");
    }

    public static String printDate(String line) {
        return findDateIn(line);
    }

    private static String findDateIn(String line) {
        String[] words = splittingLineIntoWordTable(line);
        for (String word : words) {
            if (word.matches(REGEX_IS_DATE)) {
                return word;
            }
        }
        return "";
    }

    public static String stopDate(String line) {
        return findDateIn(line);
    }

    public static Account account(String line) {
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
        return new Account(account, label.toString().trim());
    }

    public static boolean isAcccount(String line, String postalCode, String id) {
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
        String firstword = words[0];
        if (firstword.equals(postalCode)) {
            return false;
        }
        if (firstword.equals(id)) {
            return false;
        }
        if (firstword.contains("-")) {
            firstword = firstword.replace("-", "");
        }
        if (firstword.contains("4A01")) {
            firstword = firstword.replace("4A01", "401");
        }
        if (firstword.matches(REGEX_NUMBER) || "461VC".equals(firstword)) {
            return findDateIn(words[1]).isEmpty();
        }
        return false;
    }

    public static Line line(String line, Map<String, Account> accounts) {
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

        Account account;
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
        String counterpart = "";
        if (words[indexOfWords].matches(REGEX_NUMBER)) {
            counterpart = words[indexOfWords];
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
                    if (!(word.length() == 9)) {
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

        return new Line(document, date, account, journal, counterpart, checkNumber,
                label.toString().trim(), debit.toString().trim(), credit.toString().trim());
    }

    private static int getIndexOfWords(String[] words, int indexOfWords) {
        if (words[indexOfWords].trim().isEmpty()) {
            indexOfWords++;
        }
        return indexOfWords;
    }

    private static int getNumberOfAmountsOn(String line) {
        return line.split(EURO, -1).length - 1;
    }

    private static Account getAccount(Map<String, Account> accounts, String[] words, int indexOfWords) {
        Account account = null;
        if (accounts.containsKey(words[indexOfWords])) {
            account = accounts.get(words[indexOfWords]);
        } else if (words[indexOfWords].contains("-")) {
            String key = words[indexOfWords].substring(0, words[indexOfWords].indexOf("-"));
            if (key.startsWith("450")) {
                key = "45000-" + words[indexOfWords].substring(words[indexOfWords].indexOf("-") + 1);
                account = accounts.get(key);
            } else {
                account = accounts.get(key);
            }
        }
        return account;
    }

    private static String fixedSpacesBeforeEuroSign(String line) {
        line = line.replace(" €", "€");
        return line;
    }

    private static String[] splittingLineIntoWordTable(String line) {
        String[] words = line.split(" ");
        String[] resultTemp = new String[words.length + 1];
        int idResult = 0;
        for (String word : words) {
            if ("".equals(word)) {
                if (idResult > 0 && resultTemp[idResult - 1].contains(" ")) {
                    resultTemp[idResult - 1] = resultTemp[idResult - 1] + " ";
                } else {
                    resultTemp[idResult] = " ";
                    idResult++;
                }
            } else {
                resultTemp[idResult] = word;
                idResult++;
            }
        }
        if (line.endsWith(" ")) {
            resultTemp[idResult] = " ";
            idResult++;
        }
        String[] result = new String[idResult];
        System.arraycopy(resultTemp, 0, result, 0, idResult);
        return result;
    }

    private static String getDocument(String[] words, int indexOfWords) {
        // Extraction du numéro de pièce
        String document = "";
        if (findDateIn(words[indexOfWords]).isEmpty()) {
            if (!words[indexOfWords].matches(REGEX_PHONE_NUMBER)) {
                document = words[indexOfWords];
                if (document.length() == 6) {
                    document = document.substring(1, 6);
                }
            }
        }
        return document;
    }

    private static String removesStrayCharactersInLine(String line) {
        line = line.replace(" | ", " ");
        line = line.replace("| ", " ");
        line = line.replace("|", "");
        line = line.replace(" / ", " ");
        line = line.replace("Reportde", "Report de");
        line = line.replace(" ‘", "");
        line = line.replace(" — ", " ");
        line = line.replace(" . ", " ");
        line = line.replace(" =— ", " ");
        line = line.replace(" _ ", " ");
        line = line.replace(" - ", " ");
        line = line.replace(" = ", " ");
        if (line.endsWith("l")) {
            line = line.replace("l", "");
        }
        return line;
    }

    public static boolean isLigne(String line) {
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

    public static TotalAccount totalAccount(String line, Map<String, Account> accounts) {
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
        Account account = null;
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

    public static boolean isTotalAccount(String line) {
        return line.startsWith("Total compte ");
    }

    public static String postalCode(String line) {
        return line.split(" ")[0];
    }

    public static boolean isTotalBuilding(String line) {
        return line.startsWith("Total immeuble");
    }

    public static TotalBuilding totalBuilding(String line) {
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
        return new TotalBuilding(debit.toString().trim().replace(" ","").replace(EURO,""), credit.toString().trim().replace(" ","").replace(EURO,""));
    }
}
