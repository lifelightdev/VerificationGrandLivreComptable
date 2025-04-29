import org.apache.logging.log4j.LogManager;

import java.util.Map;

public class ExtractInfo {

    private static final org.apache.logging.log4j.Logger LOGGER = LogManager.getLogger();
    public static final int NAME_JOURNAL_SIZE = 2;
    public static final String EURO = "€";
    public static final String REGEX_IS_DATE = "^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(19|20)\\d{2}$";
    public static final String REGEX_NUMBER = "^-?[0-9]+$";

    public static String syndicName(String ligne) {
        return ligne.replace("|", "");
    }

    public static String printDate(String line) {
        return findDateIn(line);
    }

    private static String findDateIn(String line) {
        String[] words = line.trim().split(" ");
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
        line = line.replace(" _ "," ");
        String[] words = line.trim().split(" ");
        String account = words[0];
        StringBuilder label = new StringBuilder();
        for (String word : words) {
            if (!account.equals(word)) {
                label.append(" ").append(word);
            }
        }
        return new Account(account, label.toString().trim());
    }

    public static boolean isAcccount(String line, String postalCode, String id) {
        String[] words = line.trim().split(" ");
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
        if (firstword.matches(REGEX_NUMBER)) {
            return findDateIn(line).isEmpty();
        }
        return false;
    }

    public static Line line(String line, Map<String, Account> accounts) {
        // Supprime les caractères parasites
        line = removesStrayCharactersInLine(line);

        // Correction des espaces avant le signe €
        line = line.replace(" €","€");

        // Calcul du nombre de montant sur la ligne grace au signe €
        int nbMonaySign = line.split(EURO, -1).length - 1;

        // Découpage de la ligne en un tableau de mot
        String[] words = line.trim().split(" ");

        int indexOfWords = 0;

        String document = getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }

        // Extraction de la date de l'opération
        String date = words[indexOfWords];
        indexOfWords++;

        // Extraction du numéro de compte
        Account account;
        if (accounts.containsKey(words[indexOfWords])) {
            account = accounts.get(words[indexOfWords]);
            indexOfWords++;
        } else {
            LOGGER.error("Erreur lors de la recherche du compte {}", words[indexOfWords]);
            return null;
        }

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
        if (nbMonaySign == 3) {
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
        } else if (nbMonaySign == 1) {
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

            if (label.toString().endsWith(" ")){
                credit =amount;
            } else {
                debit =amount;
            }
        }

        return new Line(document, date, account, journal, counterpart, checkNumber,
                label.toString().trim(), debit.toString().trim(), credit.toString().trim());
    }

    private static String getDocument(String[] words, int indexOfWords) {
        // Extraction du numéro de pièce
        String document = "";
        if (findDateIn(words[indexOfWords]).isEmpty()) {
            document = words[indexOfWords];
            if (document.length() == 6)
                document = document.substring(1, 6);
        }
        return document;
    }

    private static String removesStrayCharactersInLine(String line) {
        line = line.replace(" | "," ");
        line = line.replace("| "," ");
        line = line.replace("|","");
        line = line.replace(" / "," ");
        return line;
    }

    public static boolean isLigne(String line) {
        line = removesStrayCharactersInLine(line);
        // Découpage de la ligne en un tableau de mot
        String[] words = line.trim().split(" ");
        int indexOfWords = 0;
        // Vérification que la ligne commence par une pièce
        String document = getDocument(words, indexOfWords);
        if (!document.isEmpty()) {
            indexOfWords++;
        }
        /* Vérification que la ligne commence par une date ou qu'il y a une date après la pièce */
        return !findDateIn(words[indexOfWords]).isEmpty();
    }

    public static TotalAccount totalAccount(String line) {
        return new TotalAccount("Total compte 40100-0001 (Solde : 0.00 €)", new Account("40100-0001", "label"), "100 000.00€", "100 000.00€");
    }
}
