package life.light.extract.info;

import life.light.type.TypeAccount;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Map;

public class OutilInfo {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final String EURO = "€";
    public static final String REGEX_PHONE_NUMBER = "^0[1-9]([-. ]?\\d{2}){4}$";

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
        return !words.endsWith(EURO);
    }

    public int getNumberOfAmountsOn(String line) {
        return line.split(EURO, -1).length - 1;
    }

    public TypeAccount getAccount(Map<String, TypeAccount> accounts, String[] words, int indexOfWords) {
        TypeAccount account = null;
        String key = words[indexOfWords];
        if (accounts.containsKey(key)) {
            account = accounts.get(key);
        } else if (words[indexOfWords].contains("-")) {
            key = words[indexOfWords].substring(0, words[indexOfWords].indexOf("-"));
            if (key.startsWith("450")) {
                key = "45000-" + words[indexOfWords].substring(words[indexOfWords].indexOf("-") + 1);
                account = accounts.get(key);
            } else {
                account = accounts.get(key);
            }
        } else if (key.startsWith("450")) {
            if (accounts.containsKey(key)) {
                account = accounts.get(key);
            } else {
                account = new TypeAccount(key, "Compte de tous les copropriétaires");
            }
        } else if (key.startsWith("40800")) {
            account = accounts.get("40800");
        } else if (key.startsWith("42100")) {
            account = accounts.get("42100");
        } else if (key.startsWith("43100")) {
            account = accounts.get("43100");
        } else if (key.startsWith("43200")) {
            account = accounts.get("43200");
        } else if (key.length() == 9) {
            key = key.substring(0, 5) + "-" + key.substring(5);
            account = accounts.get(key);
        }
        return account;
    }

    public String fixedSpacesBeforeEuroSign(String line) {
        line = line.replace(" €", "€");
        return line;
    }

    public String[] splittingLineIntoWordTable(String line) {
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

    public String findDateIn(String line) {
        String[] words = splittingLineIntoWordTable(line);
        String date = "";
        for (String word : words) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy", Locale.FRANCE);
            try {
                LocalDate.parse(word, formatter);
                date = word;
            } catch (Exception _) {
            }
            if (date.isEmpty()) {
                formatter = DateTimeFormatter.ofPattern("dd.MM.yy", Locale.FRANCE);
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
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        return numberOfLineInFile;
    }
}