package life.light.extract.info;

import life.light.Constant;
import life.light.type.BankLine;
import life.light.type.TypeAccount;

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

public class Bank {

    static final String DATA_FORMAT = "dd/MM/yyyy";
    public static final String ACCOUNT_BANK = "512";
    public static final String NOUVEAU_SOLDE_AU = "Nouveau solde au";
    private final OutilInfo outilInfo = new OutilInfo();
    private final Constant constant = new Constant();

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
                    try (BufferedReader reader = new BufferedReader(new FileReader(fichier))) {
                        String line;
                        TypeAccount account = accounts.get(accountBank);
                        LocalDate operationDate;
                        String label;
                        LocalDate valueDate;
                        double debit = 0D;
                        double credit;
                        while ((line = reader.readLine()) != null) {
                            String[] ligne = line.split(";");
                            if (!outilInfo.findDateIn(ligne[0]).isEmpty() && (!line.contains(NOUVEAU_SOLDE_AU))) {
                                operationDate = getOperationDate(ligne);
                                valueDate = getValueDate(ligne);
                                label = ligne[2];
                                if (line.length() > 3) {
                                    if (ligne[3].isEmpty()) {
                                        debit = 0D;
                                    } else {
                                        debit = getAmount(ligne[3]);
                                    }
                                }
                                if (ligne.length > 4) {
                                    credit = getAmount(ligne[4]);
                                } else {
                                    credit = 0D;
                                }
                                BankLine bankLine = new BankLine(year, operationDate.getMonthValue(), operationDate, valueDate, account, label.trim(), debit, credit);
                                bankLines.add(bankLine);
                            }
                        }
                    } catch (IOException e) {
                        constant.logError(Constant.LECTURE_FICHIER, e.getMessage());
                    }
                }
            }
        }
        return bankLines;
    }

    private static double getAmount(String amount) {
        return Double.parseDouble(amount.replace(".", "").replace(",", ".").replace(" ", ""));
    }

    // TODO a améliorer
    private LocalDate getOperationDate(String[] ligne) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATA_FORMAT, Locale.FRANCE);
        return LocalDate.parse(ligne[0], formatter);
    }

    // TODO a améliorer
    private LocalDate getValueDate(String[] ligne) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATA_FORMAT, Locale.FRANCE);
        return LocalDate.parse(ligne[1], formatter);
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
