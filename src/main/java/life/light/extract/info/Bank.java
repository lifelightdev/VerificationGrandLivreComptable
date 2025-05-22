package life.light.extract.info;

import life.light.type.BankLine;
import life.light.type.TypeAccount;
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

import static life.light.extract.info.OutilInfo.findDateIn;

public class Bank {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final String BANK_1_DATA_FORMAT = "dd.MM.yy";
    public static final String BANK_2_DATA_FORMAT = "dd/MM/yyyy";
    public static final String ACCOUNT_BANK = "512";
    private static final String BANK_1_ACCOUNT = "";
    private static final String BANK_2_ACCOUNT = "";

    public List<BankLine> getBankLines(Map<String, TypeAccount> accounts, String pathDirectoryBank, List<String> accountsbank, int year) {
        List<BankLine> bankLines = new ArrayList<>();
        for (String accountBank : accountsbank) {
            bankLines.addAll(getBank(accounts, pathDirectoryBank, accountBank, year));
        }
        return bankLines;
    }

    private List<BankLine> getBank(Map<String, TypeAccount> accounts, String pathDirectoryBank, String accountBank, int year) {
        List<BankLine> bankLines = new ArrayList<>();
        File pathDirectory = new File(pathDirectoryBank + File.separator + accountBank + File.separator);
        File[] files;
        files = pathDirectory.listFiles();
        if (null != files) {
            for (File fichier : files) {
                if (fichier.isFile()) {
                    //LOGGER.info("Fichier {}", fichier.getName());
                    try (BufferedReader reader = new BufferedReader(new FileReader(fichier))) {
                        String line;
                        for (int i = 0; i < 6; i++) {
                            reader.readLine();
                        }
                        TypeAccount account = accounts.get(accountBank);
                        LocalDate operationDate = null;
                        StringBuilder label = new StringBuilder();
                        LocalDate valueDate = null;
                        double debit = 0D;
                        double credit = 0D;
                        while ((line = reader.readLine()) != null) {
                            String nouveauSoldeAu = "Nouveau solde au";
                            if (!findDateIn(line).isEmpty() && (!line.contains(nouveauSoldeAu))) {
                                if (operationDate != null && !label.toString().isEmpty() && valueDate != null) {
                                    BankLine bankLine = new BankLine(year, operationDate.getMonthValue(), operationDate, valueDate, account, label.toString().trim(), debit, credit);
                                    bankLines.add(bankLine);
                                } else {
                                    String[] ligne = line.split(" ");
                                    int index = 0;
                                    operationDate = getOperationDate(accountBank, ligne, index);
                                    index = getIndexNotWord(index, ligne);
                                    if (BANK_1_ACCOUNT.equals(accountBank)) {
                                        while (!ligne[index].isEmpty()) {
                                            label.append(" ").append(ligne[index]);
                                            index++;
                                        }
                                        index = getIndexNotWord(index, ligne);
                                        valueDate = getValueDate(accountBank, ligne, index);
                                    }
                                    if (BANK_2_ACCOUNT.equals(accountBank)) {
                                        valueDate = getValueDate(accountBank, ligne, index);
                                        index = getIndexNotWord(index, ligne);
                                        while (!ligne[index].isEmpty()) {
                                            label.append(" ").append(ligne[index]);
                                            index++;
                                        }
                                    }
                                    if (line.endsWith(" ")) {
                                        debit = Double.parseDouble(ligne[ligne.length - 1].replace(".", "").replace(",", "."));
                                    } else if (!ligne[ligne.length - 1].isEmpty()) {
                                        credit = Double.parseDouble(ligne[ligne.length - 1].replace(".", "").replace(",", "."));
                                    }
                                }
                            } else if (!line.contains(nouveauSoldeAu) && (!line.contains("Total des opérations"))) {
                                if (label.toString().endsWith(" ")) {
                                    label.append(line);
                                } else {
                                    label.append(" ").append(line);
                                }
                            } else if (line.contains(nouveauSoldeAu)) {
                                BankLine bankLine = new BankLine(year, operationDate.getMonthValue(), operationDate, valueDate, account, label.toString().trim(), debit, credit);
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

    private LocalDate getOperationDate(String theAccount, String[] ligne, int index) {
        LocalDate operationDate = null;
        if (BANK_1_ACCOUNT.equals(theAccount)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(BANK_1_DATA_FORMAT, Locale.FRANCE);
            operationDate = LocalDate.parse(ligne[index] + ".24", formatter);
        } else if (BANK_2_ACCOUNT.equals(theAccount)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(BANK_2_DATA_FORMAT, Locale.FRANCE);
            operationDate = LocalDate.parse(ligne[index], formatter);
        }
        return operationDate;
    }

    private LocalDate getValueDate(String theAccount, String[] ligne, int index) {
        LocalDate operationDate = null;
        if (BANK_1_ACCOUNT.equals(theAccount)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(BANK_1_DATA_FORMAT, Locale.FRANCE);
            operationDate = LocalDate.parse(ligne[index], formatter);
        } else if (BANK_2_ACCOUNT.equals(theAccount)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(BANK_2_DATA_FORMAT, Locale.FRANCE);
            operationDate = LocalDate.parse(ligne[index], formatter);
        }
        return operationDate;
    }

    private int getIndexNotWord(int index, String[] ligne) {
        do {
            index++;
        } while (ligne[index].isEmpty());
        return index;
    }

    public static List<TypeAccount> getBanks(Map<String, TypeAccount> accounts) {
        List<TypeAccount> accountsBank = new ArrayList<>();
        for (TypeAccount account : accounts.values()) {
            if (account.account().startsWith(ACCOUNT_BANK)) {
                accountsBank.add(account);
            }
        }
        return accountsBank;
    }
}
