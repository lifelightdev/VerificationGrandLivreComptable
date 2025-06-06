package life.light.extract.info;

import life.light.Constant;
import life.light.type.*;
import life.light.write.WriteOutil;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;

import static life.light.Constant.DATE_FORMATTER;
import static life.light.Constant.TOTAL_IMMEUBLE;

public class ExpenseExtract {

    private static final String CLE = "Clé";
    private static final String CODE_NATURE = "Code Nature ";
    private static final String TOTAL_NATURE = "Total Nature ";
    private static final String TOTAL_CLE = "Total clé ";

    private final OutilInfo outilInfo = new OutilInfo();
    private final WriteOutil outilWrite = new WriteOutil();
    private final Constant constant = new Constant();

    public LineOfExpense[] getList(String fileName) {
        LineOfExpense[] listOfExpense = new LineOfExpense[outilInfo.getNumberOfLineInFile(fileName)];
        int indexLine = 0;
        OutilInfo outilInfo = new OutilInfo();
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            for (int i = 0; i < 8; i++) {
                reader.readLine();
            }
            while ((line = reader.readLine()) != null) {
                if (line.startsWith(CLE)) {
                    listOfExpense[indexLine++] = getLineOfExpenseKey(outilInfo, line, TypeOfExpense.Key);
                } else if (line.startsWith(CODE_NATURE)) {
                    listOfExpense[indexLine++] = getLineOfExpenseKey(outilInfo, line, TypeOfExpense.Nature);
                } else if (line.startsWith(TOTAL_NATURE)) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Nature);
                } else if (line.startsWith(TOTAL_CLE)) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Key);
                } else if (line.startsWith(TOTAL_IMMEUBLE)) {
                    listOfExpense[indexLine++] = getLineOfExpenseTotal(outilInfo, line, TypeOfExpense.Building);
                } else if (!outilInfo.findDateIn(line).isEmpty() && !line.contains("Liste des dépenses")) {
                    String[] words = outilInfo.getWords(line);
                    int index = outilInfo.getIndexOfNextWords(words, 0);
                    String document = words[index++];
                    index = outilInfo.getIndexOfNextWords(words, index);
                    LocalDate date = LocalDate.parse(words[index++], DATE_FORMATTER);
                    index = outilInfo.getIndexOfNextWords(words, index);

                    // Calcule de nombre de montants sur la ligne grâce au signe €
                    long numberOfAmount = outilInfo.getNumberOfAmountsOn(line);

                    StringBuilder label = new StringBuilder(words[index]);
                    StringBuilder amount = new StringBuilder();
                    String recovery = "";
                    String deduction = "";
                    int i = words.length - 2;
                    if (numberOfAmount == 3) {
                        WordInWords wordInWords = getAmoutInWords(words, i);
                        recovery = wordInWords.word();
                        wordInWords = getAmoutInWords(words, wordInWords.index());
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
                    if (numberOfAmount == 2) {
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

                    LineOfExpenseValue lineOfExpense = new LineOfExpenseValue(document, date, label.toString().trim(),
                            amount.toString(), deduction, recovery, null, null);
                    listOfExpense[indexLine++] = lineOfExpense;
                }

            }
        } catch (IOException e) {
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
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

    public int getAccountingYear(String pathFileListOfExpenses) {
        try (BufferedReader reader = new BufferedReader(new FileReader(pathFileListOfExpenses))) {
            outilInfo.readNextLineInFile(reader, 4);
            String line = reader.readLine();
            String[] words = outilInfo.getWords(line.trim());
            return Integer.parseInt(words[words.length - 1]);
        } catch (IOException e) {
            constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
        }
        return 0;
    }

    private record WordInWords(String word, int index) {
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
