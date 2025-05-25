package life.light.extract.info;

import life.light.type.LineOfExpense;
import life.light.type.LineOfExpenseKey;
import life.light.type.LineOfExpenseTotal;
import life.light.type.TypeOfExpense;
import life.light.write.OutilWrite;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;

import static life.light.write.OutilWrite.DATE_FORMATTER;

public class Expense {

    private static final Logger LOGGER = LogManager.getLogger();

    OutilInfo outilInfo = new OutilInfo();
    OutilWrite outilWrite = new OutilWrite();

    public Object[] getList(String fileName) {
        Object[] listOfExpense = new Object[outilInfo.getNumberOfLineInFile(fileName)];
        int indexLine = 0;
        OutilInfo outilInfo = new OutilInfo();
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            for (int i = 0; i < 8; i++) {
                reader.readLine();
            }
            while ((line = reader.readLine()) != null) {
                if (line.startsWith("Clé")) {
                    listOfExpense[indexLine++] = getLineOfExpenseKey(outilInfo, line, TypeOfExpense.Key);
                } else if (line.startsWith("Code Nature ")) {
                    listOfExpense[indexLine++] = getLineOfExpenseKey(outilInfo, line, TypeOfExpense.Nature);
                } else if (line.startsWith("Total Nature ")) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Nature);
                } else if (line.startsWith("Total clé : ")) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Key);
                } else if (line.startsWith("Total immeuble : ")) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Building);
                } else if (!outilInfo.findDateIn(line).isEmpty() && !line.contains("ste des dépense")) {
                    String[] words = outilInfo.getWords(line);
                    int index = outilInfo.getIndexOfNextWords(words, 0);
                    String document = words[index++];
                    index = outilInfo.getIndexOfNextWords(words, index);
                    LocalDate date = LocalDate.parse(words[index++], DATE_FORMATTER);
                    index = outilInfo.getIndexOfNextWords(words, index);

                    long numberOfAmout = line.chars()
                            .filter(c -> c == '€')
                            .count();

                    StringBuilder label = new StringBuilder(words[index]);
                    StringBuilder amount = new StringBuilder();
                    String recovery = "";
                    String deduction = "";
                    int i = words.length - 2;
                    if (numberOfAmout == 3) {
                        WordInWords wordInWords = getAmoutInWords(words, i);
                        recovery = wordInWords.word();
                        i = wordInWords.index();
                        wordInWords = getAmoutInWords(words, i);
                        deduction = wordInWords.word();
                        i = wordInWords.index();

                        while (outilWrite.isDouble(words[i])) {
                            amount.insert(0, words[i]);
                            i--;
                        }
                        while (index < i) {
                            label.append(" ").append(words[index++]);
                        }
                    }
                    if (numberOfAmout == 2) {
                        if (line.endsWith(" ")) {
                            WordInWords wordInWords = getAmoutInWords(words, i);
                            deduction = wordInWords.word();
                            i = wordInWords.index();
                        } else {
                            WordInWords wordInWords = getAmoutInWords(words, i);
                            recovery = wordInWords.word();
                            i = wordInWords.index();
                        }
                        while (outilWrite.isDouble(words[i])) {
                            amount.insert(0, words[i]);
                            i--;
                        }
                        while (index < i) {
                            label.append(" ").append(words[index++]);
                        }
                    }

                    LineOfExpense lineOfExpense = new LineOfExpense(document, date, label.toString().trim(), amount.toString(), deduction, recovery);
                    listOfExpense[indexLine++] = lineOfExpense;
                }

            }
        } catch (IOException e) {
            LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
        }
        return listOfExpense;
    }

    private LineOfExpenseKey getLineOfExpenseKey(OutilInfo outilInfo, String line, TypeOfExpense type) {
        line = line.replace(":", "").trim();
        String[] words = outilInfo.getWords(line);
        int index = 0;
        WordInWords wordInWords = getWordInWords(words, index);
        String label = wordInWords.word();
        index = wordInWords.index();
        String key = words[index++];
        String value = getWordInWords(words, index).word();
        return new LineOfExpenseKey(key, label.trim(), value.trim(), type);
    }

    private LineOfExpenseTotal getLineOfExpenseTotal(OutilInfo outilInfo, String line, TypeOfExpense type) {
        line = line.replace(":", "").trim();
        String[] words = outilInfo.getWords(line);
        int index = 0;
        WordInWords wordInWords = getWordInWords(words, index);
        String label = wordInWords.word().trim();
        index = wordInWords.index();
        String key = words[index++];
        StringBuilder value = new StringBuilder(getWordInWords(words, index).word());
        StringBuilder amount = new StringBuilder();
        String recovery;
        String deduction;
        int i = words.length - 2;
        wordInWords = getAmoutInWords(words, i);
        recovery = wordInWords.word();
        i = wordInWords.index();
        wordInWords = getAmoutInWords(words, i);
        deduction = wordInWords.word();
        i = wordInWords.index();
        while (outilWrite.isDouble(words[i])) {
            amount.insert(0, words[i]);
            i--;
        }
        while (index < i) {
            value.append(" ").append(words[index++]);
        }
        return new LineOfExpenseTotal(key, label.trim(), value.toString().trim(), amount.toString(), deduction, recovery, type);
    }

    record WordInWords(String word, int index) {
    }

    private WordInWords getWordInWords(String[] words, int index) {
        StringBuilder word = new StringBuilder();
        while (index < words.length && !outilWrite.isDouble(words[index])) {
            word.append(" ").append(words[index++]);
        }
        return new WordInWords(word.toString(), index);
    }

    private WordInWords getAmoutInWords(String[] words, int index) {
        StringBuilder word = new StringBuilder();
        while (outilInfo.isAmount(words[index])) {
            word.insert(0, words[index]);
            index--;
        }
        index--;
        return new WordInWords(word.toString(), index);
    }
}
