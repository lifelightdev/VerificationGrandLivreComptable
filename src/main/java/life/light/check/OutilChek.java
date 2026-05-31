package life.light.check;

import life.light.type.BankLine;
import life.light.type.LineLedger;
import life.light.type.LineOfStateOfReconciliation;
import life.light.type.StateOfReconciliation;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class OutilChek {
    public File getFileInDirectory(String document, File file) {
        File fileFound = null;
        File[] listOfFiles = file.listFiles();
        if (null != listOfFiles) {
            for (File fileOfDirectory : listOfFiles) {
                if (fileOfDirectory.isFile()) {
                    if (fileOfDirectory.getName().contains(document)) {
                        fileFound = fileOfDirectory;
                        break;
                    }
                } else if (fileOfDirectory.isDirectory()) {
                    if (fileFound == null) {
                        fileFound = getFileInDirectory(document, fileOfDirectory);
                    } else {
                        break;
                    }
                }
            }
        }
        return fileFound;
    }

    public static StateOfReconciliation getStateOfConvergence(List<LineLedger> ledger, List<BankLine> bankStatement) {
        List<LineOfStateOfReconciliation> find = new ArrayList<>();
        List<LineOfStateOfReconciliation> noFindInLedger = new ArrayList<>();
        List<LineOfStateOfReconciliation> noFindInBank = new ArrayList<>();
        for (LineLedger lineLedger : ledger) {
            for (BankLine bankLine : bankStatement) {
                String message;
                if (lineLedger.account().account().equals(bankLine.account().account())) {
                    if (lineLedger.amountCredit() != 0) {
                        if (lineLedger.amountCredit().equals(bankLine.debit())) {
                            if (lineLedger.amountDebit().equals(bankLine.credit())) {
                                message = "OK les deux montants sont vérifié";
                                if (lineLedger.accountCounterpart().account().startsWith("450")) {
                                    String probablyADundraisingAppeal = isProbablyADundraisingAppeal(lineLedger, bankLine);
                                    if (probablyADundraisingAppeal != null) {
                                        message = "OK les deux montants sont vérifié et " + probablyADundraisingAppeal;
                                    }
                                }
                                if (lineLedger.accountCounterpart().account().startsWith("40100")) {
                                    if (isProbablyASupplier(lineLedger, bankLine)) {
                                        message = "OK les deux montants sont vérifié et il semble que c'est le règlement d'un fournisseur";
                                    }
                                }
                                LineOfStateOfReconciliation resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, message);
                                find.add(resultLine);
                                break;
                            }
                        }
                    }
                    if (lineLedger.amountDebit() != 0) {
                        if (lineLedger.amountDebit().equals(bankLine.credit())) {
                            if (lineLedger.amountCredit().equals(bankLine.debit())) {
                                LineOfStateOfReconciliation resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, "OK les deux montants sont vérifié");
                                if (lineLedger.accountCounterpart().account().startsWith("450") && !lineLedger.accountCounterpart().account().equals("450000000")) {
                                    String probablyADundraisingAppeal = isProbablyADundraisingAppeal(lineLedger, bankLine);
                                    if (probablyADundraisingAppeal != null) {
                                        message = "OK les deux montants sont vérifié et " + probablyADundraisingAppeal;
                                        resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, message);
                                    }
                                }
                                if (lineLedger.accountCounterpart().account().equals("40100")) {
                                    if (isProbablyASupplier(lineLedger, bankLine)) {
                                        resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, "OK les deux montants sont vérifié et il semble que c'est le règlement d'un fournisseur");
                                    }
                                }
                                find.add(resultLine);
                                break;
                            }
                        }
                    }
                    // Verification du libellé
                    String[] ledgerLabels = lineLedger.label().split(" ");
                    for (String word : ledgerLabels) {
                        if (isAGoodWord(word)) {
                            if (bankLine.label().toUpperCase().contains(word.toUpperCase())) {
                                message = "KO - Il n'y a pas de correspondance stricte entre les montants. Mais il y a le même mot dans les deux libellé : " + word;
                                //System.out.println(message + " ------ " + lineLedger + " ------------- " + bankLine);
                                LineOfStateOfReconciliation resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, message);
                                noFindInLedger.add(resultLine);
                                break;
                            }
                        }
                    }
                    // Vérification du libellé du relever avec le libellé de la contrepartie dans le grand livre
                    ledgerLabels = lineLedger.accountCounterpart().label().split(" ");
                    for (String word : ledgerLabels) {
                        if (isAGoodWord(word)) {
                            if (bankLine.label().toUpperCase().contains(word.toUpperCase())) {
                                message = "KO - Il n'y a pas de correspondance stricte entre les montants. " +
                                        "Mais il y a le même mot dans le libellé de l'opération bancaire et dans le libellé du compte de contrepartie : " + word;
                                LineOfStateOfReconciliation resultLine = new LineOfStateOfReconciliation(lineLedger, bankLine, message);
                                noFindInLedger.add(resultLine);
                                break;
                            }
                        }
                    }
                }
            }
        }
        return new StateOfReconciliation(find, noFindInLedger, noFindInBank);
    }

    private static String isProbablyADundraisingAppeal(LineLedger lineLedger, BankLine bankLine) {
        String message = null;
        if (bankLine.label().contains("APPEL")) {
            return "Il semble que ce soit un appel de fond";
        }
        if (bankLine.label().contains("af")) {
            return "Il semble que ce soit un appel de fond";
        }
        String subAccount = lineLedger.accountCounterpart().account().substring(lineLedger.accountCounterpart().account().length() - 3);
        if (bankLine.label().contains(subAccount)) {
            return "Il semble que ce soit un appel de fond pour le compte 450-" + subAccount;
        }
        String subLabelAccount = lineLedger.accountCounterpart().label().substring(0, lineLedger.accountCounterpart().label().lastIndexOf("-") - 1).toUpperCase();
        if (bankLine.label().toUpperCase().contains(subLabelAccount)) {
            return "Il semble que ce soit un appel de fond du propriétaire : " + subLabelAccount;
        }
        String[] words = subLabelAccount.split(" ");
        for (String word : words) {
            if (!word.equalsIgnoreCase("ET") && !word.equalsIgnoreCase("LE") && !word.equalsIgnoreCase("OU")) {
                if (bankLine.label().toUpperCase().contains(word.toUpperCase())) {
                    return "Il semble que ce soit un appel de fond du propriétaire : " + word;
                }
            }
        }
        return message;
    }

    private static boolean isProbablyASupplier(LineLedger lineLedger, BankLine bankLine) {
        String[] words = lineLedger.accountCounterpart().label().split(" ");
        for (String word : words) {
            if (bankLine.label().toUpperCase().contains(word.toUpperCase())) {
                return true;
            }
        }
        return false;
    }

    private static boolean isAGoodWord(String word) {
        String[] badWords = {"VIREMENT", "VI", "ME", "EP", "DES", "CT", "VOTRE", "COURANT", "au", "PRLV", "PRELEVEMENT", "DE", "2025", "DA", "B",
                "en", "et", "REM", "IMMO", "APPEL", "-", "FONDS", "LE", "VIR", "A", "PRLV", "CHQ", "OU", "ou", "SI", "DU", "PREL", "Mme"};
        for (String badWord : badWords) {
            if (word.equals(badWord)) {
                return false;
            }
        }
        return true;
    }
}
