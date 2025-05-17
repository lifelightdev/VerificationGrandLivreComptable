package life.light.extract.info;

import life.light.type.TypeAccount;
import life.light.type.BankLine;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import static life.light.Main.*;
import static life.light.extract.info.OutilInfo.findDateIn;

public class Bank {

    private static final Logger LOGGER = LogManager.getLogger();

    private Bank() { }

    public static List<BankLine> getBankLines(Map<String, TypeAccount> accounts, List<String> pathsDirectoryBank) {
        List<BankLine> bankLines = new ArrayList<>();
        for (String pathDirectoryBank : pathsDirectoryBank) {
            bankLines.addAll(getBank(accounts, pathDirectoryBank));
        }
        return bankLines;
    }

    private static List<BankLine> getBank(Map<String, TypeAccount> accounts, String pathDirectoryBank) {
        List<BankLine> bankLines = new ArrayList<>();
        File pathDirectory;
        File[] files;
        String[] path = pathDirectoryBank.split("\\\\");
        String theAccount = path[path.length - 1];
        pathDirectory = new File(pathDirectoryBank);
        files = pathDirectory.listFiles();
        DateTimeFormatter formatter = null;
        if (ACCOUNT_51220.equals(theAccount)) {
            formatter = DateTimeFormatter.ofPattern("dd.MM.yy", Locale.FRANCE);
        }
        if (ACCOUNT_51221.equals(theAccount) || "512".equals(theAccount)) {
            formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy", Locale.FRANCE);
        }
        if (null != formatter && null != files) {
            for (File fichier : files) {
                if (fichier.isFile()) {
                    //LOGGER.info("Fichier {}", fichier.getName());
                    try (BufferedReader reader = new BufferedReader(new FileReader(fichier))) {
                        String line;
                        for (int i = 0; i < 6; i++) {
                            reader.readLine();
                        }
                        while ((line = reader.readLine()) != null) {
                            if (!findDateIn(line).isEmpty() && (!line.contains("Nouveau solde au"))) {
                                String[] ligne = line.split(" ");
                                int index = 0;
                                TypeAccount account = accounts.get(theAccount);
                                LocalDate operationDate = getOperationDate(theAccount, formatter, ligne, index);
                                index = getIndexNotWord(index, ligne);
                                StringBuilder label = new StringBuilder();
                                LocalDate valueDate = null;
                                if (ACCOUNT_51220.equals(theAccount)) {
                                    label = new StringBuilder(ligne[index++]);
                                    while (!ligne[index].isEmpty()) {
                                        label.append(" ").append(ligne[index]);
                                        index++;
                                    }
                                } else if (ACCOUNT_51221.equals(theAccount)|| "512".equals(theAccount)) {
                                    valueDate = LocalDate.parse(ligne[index], formatter);
                                }
                                index = getIndexNotWord(index, ligne);
                                if (ACCOUNT_51220.equals(theAccount)) {
                                    valueDate = LocalDate.parse(ligne[index], formatter);
                                } else if (ACCOUNT_51221.equals(theAccount) || "512".equals(theAccount)) {
                                    label = new StringBuilder(ligne[index++]);
                                    while (!ligne[index].isEmpty()) {
                                        label.append(" ").append(ligne[index]);
                                        index++;
                                    }
                                }

                                double debit = 0D;
                                double credit = 0D;
                                if (line.endsWith(" ")) {
                                    debit = Double.parseDouble(ligne[ligne.length - 1].replace(".", "").replace(",", "."));
                                } else if (!ligne[ligne.length - 1].isEmpty()) {
                                    credit = Double.parseDouble(ligne[ligne.length - 1].replace(".", "").replace(",", "."));
                                }
                                BankLine bankLine = new BankLine(2024, operationDate.getMonthValue(), operationDate, valueDate, account, label.toString().trim(), debit, credit);
                                //LOGGER.info("Date {} {} {} {} débit {} crédit {}", operationDate, valueDate, account, label.trim(), debit, credit);
                                bankLines.add(bankLine);
                            }
                        }
                    } catch (IOException e) {
                        LOGGER.error("Erreur lors de la lecture du fichier avec cette erreur {}", e.getMessage());
                    }
                }
            }
        }
        return bankLines;
    }

    private static LocalDate getOperationDate(String theAccount, DateTimeFormatter formatter, String[] ligne, int index) {
        LocalDate operationDate = null;
        if (ACCOUNT_51220.equals(theAccount)) {
            operationDate = LocalDate.parse(ligne[index] + ".24", formatter);
        } else if (ACCOUNT_51221.equals(theAccount) || "512".equals(theAccount)) {
            operationDate = LocalDate.parse(ligne[index], formatter);
        }
        return operationDate;
    }

    private static int getIndexNotWord(int index, String[] ligne) {
        do {
            index++;
        } while (ligne[index].isEmpty());
        return index;
    }
}
