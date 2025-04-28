import org.apache.logging.log4j.LogManager;

import java.util.Map;

public class ExtractInfo {

    private static final org.apache.logging.log4j.Logger LOGGER = LogManager.getLogger();

    public static String syndicName(String ligne) {
        return ligne.replace("|", "");
    }

    public static String printDate(String line) {
        return findDateIn(line);
    }

    private static String findDateIn(String line) {
        String[] words = line.trim().split(" ");
        for (String word : words) {
            if (word.matches("^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/(19|20)\\d{2}$")) {
                return word;
            }
        }
        return "";
    }

    public static String stopDate(String line) {
        return findDateIn(line);
    }

    public static Account account(String line) {
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
        if (firstword.matches("^-?[0-9]+$")) {
            return findDateIn(line).isEmpty();
        }
        return false;
    }

    public static Line line(String line, Map<String, Account> accounts) {

        int nbMonaySign = line.split("€", -1).length - 1;
        String[] words = line.trim().split(" ");

        String document = "";
        int indexOfWords = 0;
        if (findDateIn(words[indexOfWords]).isEmpty()) {
            document = words[indexOfWords];
            indexOfWords++;
        }

        String date = words[indexOfWords];
        indexOfWords++;

        Account account;
        if (accounts.containsKey(words[indexOfWords])) {
            account = accounts.get(words[indexOfWords]);
            indexOfWords++;
        } else {
            LOGGER.error("Erreur lors de la recherche du compte {}", words[indexOfWords]);
            return null;
        }

        String journal = "";
        if (words[indexOfWords].length() == 2) {
            journal = words[indexOfWords];
            indexOfWords++;
        }

        String counterpart = "";
        if (words[indexOfWords].length() == 5 && words[indexOfWords].matches("^-?[0-9]+$")) {
            counterpart = words[indexOfWords];
            indexOfWords++;
        }
        StringBuilder label = new StringBuilder();
        StringBuilder debit = new StringBuilder();
        StringBuilder credit = new StringBuilder();
        if (nbMonaySign == 3) {
            while (!words[indexOfWords].equals("€")) {
                label.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            label.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].equals("€")) {
                debit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            debit.append(" ").append(words[indexOfWords]);
            indexOfWords++;
            while (!words[indexOfWords].equals("€")) {
                credit.append(" ").append(words[indexOfWords]);
                indexOfWords++;
            }
            credit.append(" ").append(words[indexOfWords]);
        } else if (nbMonaySign == 1) {
            label.append("APPEL FONDS LOI ALUR");
            credit.append("2 000.00 €");
        }

        return new Line(document, date, account, journal, counterpart, label.toString().trim(), debit.toString().trim(), credit.toString().trim());
    }
}
