package life.light.extract.info;

import life.light.Constant;
import life.light.type.TypeAccount;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Map;

import static life.light.Constant.*;

public class OutilInfo {

    public static final String REGEX_PHONE_NUMBER = "^0[1-9]([-. ]?\\d{2}){4}$";
    public static final String ACCOUNT_CO_OWNER = "450";
    private final Constant constant = new Constant();

    public int getIndexOfNextWords(String[] words, int indexOfLastWords) {
        if (words[indexOfLastWords].trim().isEmpty()) {
            indexOfLastWords++;
        }
        return indexOfLastWords;
    }

    public String[] getWords(String line) {
        return line.split(" ");
    }

    public boolean isAmount(String words) {
        return !words.endsWith(Character.toString(EURO));
    }

    public long getNumberOfAmountsOn(String line) {
        return line.chars()
                .filter(c -> c == EURO)
                .count();
    }

    public TypeAccount getAccount(Map<String, TypeAccount> accounts, String word) {
        TypeAccount account = null;
        String key = word;
        if (accounts.containsKey(key)) {
            account = accounts.get(key);
        } else if (key.startsWith("40800")) {
            account = accounts.get("40800");
        } else if (key.startsWith("42100")) {
            account = accounts.get("42100");
        } else if (key.startsWith("43100")) {
            account = accounts.get("43100");
        } else if (key.startsWith("43200")) {
            account = accounts.get("43200");
        } else if (key.startsWith("44200")) {
            account = accounts.get("44200");
        } else if (key.startsWith("45020")) {
            if (word.contains("-")) {
                account = accounts.get("45000" + key.substring(key.indexOf("-")));
                String label = account.label() + " -- " + accounts.get("10241").label();
                account = new TypeAccount(key, label);
            } else {
                account = accounts.get(key);
            }
        } else if (key.startsWith("45050")) {
            if (word.contains("-")) {
                account = accounts.get("45000" + key.substring(key.indexOf("-")));
                String label = account.label() + " -- " + accounts.get("10500").label();
                account = new TypeAccount(key, label);
            } else {
                account = accounts.get(key);
            }
        } else if (key.startsWith("45010")) {
            if (word.contains("-")) {
                account = accounts.get("45000" + key.substring(key.indexOf("-")));
                String label = account.label() + " -- " + accounts.get("70120").label();
                account = new TypeAccount(key, label);
            } else {
                if (key.length() > 5) {
                    account = accounts.get("45000-" + key.substring(5));
                    String label = account.label() + " -- " + accounts.get("70120").label();
                    account = new TypeAccount(key, label);
                } else {
                    account = accounts.get(key);
                }
            }
        } else if (key.startsWith("45030")) {
            account = new TypeAccount(key, "Inconu");
        } else if (key.startsWith("450")) {
            account = accounts.get("45000" + key.substring(key.indexOf("-")));
        } else if (word.contains("-")) {
            key = word.replace("-", "").toLowerCase();
            account = accounts.get(key);
        } else if (key.startsWith("40100") && !key.contains("-")) {
            account = accounts.get("40100-" + key.substring(5));
        }
        return account;
    }

    public String fixedSpacesBeforeEuroSign(String line) {
        line = line.replace(" " + EURO, Character.toString(EURO));
        return line;
    }

    public String[] splittingLineIntoWordTable(String line) {
        return line.split(";");
    }

    public String getDocument(String[] words, int indexOfWords) {
        // Extraction du numéro de pièce
        String document = "";
        if (findDateIn(words[indexOfWords]).isEmpty() && !words[indexOfWords].matches(REGEX_PHONE_NUMBER)) {
            document = words[indexOfWords];
            if (document.length() == 6) {
                document = document.substring(1, 6);
            }
        }
        return document;
    }

    public String removesStrayCharactersInLine(String line) {
        line = line.replace(" | ", " ");
        line = line.replace("| ", " ");
        line = line.replace("|", "");
        line = line.replace(" / ", " ");
        line = line.replace("Reportde", REPORT_DE);
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

    public String findDateIn(String line) {
        String[] words = splittingLineIntoWordTable(line);
        String date = "";
        for (String word : words) {
            try {
                LocalDate.parse(word, DATE_FORMATTER);
                date = word;
            } catch (Exception _) {
            }
            if (date.isEmpty()) {
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yy", Locale.FRANCE);
                try {
                    LocalDate.parse(word, formatter);
                    date = word;
                } catch (Exception _) {
                }
            }
        }
        return date;
    }

    public int getNumberOfLineInFile(String path) {
        int numberOfLineInFile = 0;
        try (BufferedReader reader = new BufferedReader(new FileReader(path))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (!line.trim().isEmpty()) {
                    numberOfLineInFile++;
                }
            }
        } catch (IOException e) {
            constant.logError(LECTURE_FICHIER, e.getMessage());
        }
        return numberOfLineInFile;
    }

    public void readNextLineInFile(BufferedReader reader, int endLine) throws IOException {
        for (int i = 0; i < endLine; i++) {
            reader.readLine();
        }
    }
}