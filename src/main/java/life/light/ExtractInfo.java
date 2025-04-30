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
        if (firstword.matches(REGEX_NUMBER) || "461VC".equals(firstword)) {
            return findDateIn(words[1]).isEmpty() && !line.contains("Page : ");
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

        // Extraction de la date de l'opération
        String date = words[indexOfWords];
        indexOfWords++;

        // Extraction du numéro de compte
        Account account = getAccount(accounts, words, indexOfWords);
        if (account == null) {
            LOGGER.error("Erreur lors de la recherche du compte {} sur la ligne {}", words[indexOfWords], line);
            return null;
        }
        indexOfWords++;

        // Extraction du code du journal
        String journal = "";
        if (words[indexOfWords].length() == NAME_JOURNAL_SIZE) {
            journal = words[indexOfWords];
            indexOfWords++;
        }

        // Extraction du compte de contrepartie
        String counterpart = "";
        if (words[indexOfWords].length() == 5 && words[indexOfWords].matches(REGEX_NUMBER)) {
            counterpart = words[indexOfWords];
            indexOfWords++;
        }

        // Extraction du numéro de chéque
        String checkNumber = "";
        if ("Virt".equals(words[indexOfWords])) {
            checkNumber = words[indexOfWords];
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
            // S'il n'y a qu'un seul montant
            // c'est ne nombre d'espace entre le libellé et le montant qui determine s'il est au debit ou au crédit
            // zero espace = debit
            // un espace = crédit
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
        }

        if (!debit.toString().contains(".") && !debit.toString().isEmpty()) {
            debit = new StringBuilder(debit.substring(0, debit.toString().length() - 3) + "." + debit.substring(debit.toString().length() - 3));
        }

        debit = new StringBuilder(debit.toString().replace(" ", "").replace("€", "").trim());
        credit = new StringBuilder(credit.toString().replace(" ", "").replace("€", "").trim());

        return new Line(document, date, account, journal, counterpart, checkNumber,
                label.toString().trim(), debit.toString().trim(), credit.toString().trim());
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
        return line.trim().split(" ");
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
        String[] labels = splittingLineIntoWordTable(label.toString());
        Account account = null;
        for (int indexOfLabel = 0; indexOfLabel < labels.length; indexOfLabel++) {
            account = getAccount(accounts, labels, indexOfLabel);
            if (account != null) {
                label = new StringBuilder(label.toString().replace(account.account(), "").replace("  ", " "));
                break;
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
}
