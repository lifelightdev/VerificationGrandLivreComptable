package life.light.check;

import life.light.extract.info.ExpenseExtract;
import life.light.type.*;
import life.light.write.WriteFile;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeMap;

import static life.light.Constant.PATH;
import static life.light.write.WriteOutil.IMPOSSIBLE_DE_TROUVER_LA_PIECE;

public class Expense {
    private final OutilChek outilChek = new OutilChek();

    public LineOfExpense[] getListOfExpense(String filePath) {
        ExpenseExtract expense = new ExpenseExtract();
        return expense.getList(filePath);
    }

    public List<String> checkTotal(LineOfExpense[] listOfExpense) {
        String codeKey = "";
        String codeNature = "";
        double totalNatureAmount = 0D;
        double totalNatureDeduction = 0D;
        double totalNatureRecovery = 0D;
        double totalKeyAmount = 0D;
        double totalKeyDeduction = 0D;
        double totalKeyRecovery = 0D;
        double totalBuildingAmount = 0D;
        double totalBuildingDeduction = 0D;
        double totalBuildingRecovery = 0D;
        String message = "";
        List<String> listOfMessage = new ArrayList<>();
        for (LineOfExpense lineOfExpense : listOfExpense) {
            if (lineOfExpense instanceof LineOfExpenseKey) {
                if (((LineOfExpenseKey) lineOfExpense).type().equals(TypeOfExpense.Key)) {
                    codeKey = ((LineOfExpenseKey) lineOfExpense).label();
                } else if (((LineOfExpenseKey) lineOfExpense).type().equals(TypeOfExpense.Nature)) {
                    codeNature = ((LineOfExpenseKey) lineOfExpense).label();
                }
            } else if (lineOfExpense instanceof LineOfExpenseValue line) {
                if (((line.keyCode()).equals(codeKey))
                        && (line.natureCode().equals(codeNature))) {
                    totalNatureAmount += Double.parseDouble(line.amount());
                    totalNatureDeduction += Double.parseDouble(line.deduction());
                    totalNatureRecovery += Double.parseDouble(line.recovery());
                }
            } else if (lineOfExpense instanceof LineOfExpenseTotal line) {
                if ((line.type().equals(TypeOfExpense.Nature))
                        && (line.key().equals(codeNature))) {
                    totalKeyAmount += totalNatureAmount;
                    if (totalNatureAmount != Double.parseDouble(line.amount())) {
                        message = "Pour le code nature " + codeNature + " : \n" +
                                "Le total du montant n'est pas correcte. "
                                + "Le total écrit est " + line.amount() + " alors que le calcul donne " + totalNatureAmount + "\n";
                    }
                    totalNatureAmount = 0D;
                    totalKeyDeduction += totalNatureDeduction;
                    if (totalNatureDeduction != Double.parseDouble(line.deduction())) {
                        if (message.isEmpty()) {
                            message = "Pour le code nature " + codeNature + " : \n";
                        }
                        message += "Le total de la déduction n'est pas correcte. Le total écrit est " + line.deduction()
                                + " alors que le calcul donne " + totalNatureDeduction + "\n";
                    }
                    totalNatureDeduction = 0D;
                    totalKeyRecovery += totalNatureRecovery;
                    if (totalNatureRecovery != Double.parseDouble(line.recovery())) {
                        if (message.isEmpty()) {
                            message = "Pour le code nature " + codeNature + " : \n";
                        }
                        message += "Le total de la récuperation n'est pas correcte. Le total écrit est " + line.recovery()
                                + " alors que le calcul donne " + totalNatureRecovery + "\n";
                    }
                    totalNatureRecovery = 0D;
                    if (!message.isEmpty()) {
                        listOfMessage.add(message);
                    }
                } else if ((line.type().equals(TypeOfExpense.Key))
                        && (line.key().equals(codeKey))) {
                    if (totalKeyAmount != Double.parseDouble(line.amount())) {
                        message = "Pour le code clé " + codeKey + " : \n" +
                                "Le total du montant n'est pas correcte. "
                                + "Le total écrit est " + line.amount() + " alors que le calcul donne " + totalKeyAmount + "\n";
                    }
                    totalBuildingAmount += totalKeyAmount;
                    totalKeyAmount = 0D;
                    if (totalKeyDeduction != Double.parseDouble(line.deduction())) {
                        if (message.isEmpty()) {
                            message = "Pour le code clé " + codeKey + " : \n";
                        }
                        message += "Le total de la déduction n'est pas correcte. Le total écrit est " + line.deduction()
                                + " alors que le calcul donne " + totalKeyDeduction + "\n";
                    }
                    totalBuildingDeduction += totalKeyDeduction;
                    totalKeyDeduction = 0D;
                    if (totalKeyRecovery != Double.parseDouble(line.recovery())) {
                        if (message.isEmpty()) {
                            message = "Pour le code clé " + codeKey + " : \n";
                        }
                        message += "Le total de la récuperation n'est pas correcte. Le total écrit est " + line.recovery()
                                + " alors que le calcul donne " + totalKeyRecovery + "\n";
                    }
                    totalBuildingRecovery += totalKeyRecovery;
                    totalKeyRecovery = 0D;
                    if (!message.isEmpty()) {
                        listOfMessage.add(message);
                    }
                } else if ((line.type().equals(TypeOfExpense.Building))) {
                    if (totalBuildingAmount != Double.parseDouble(line.amount())) {
                        message = "Pour l'immeuble :\n" +
                                "Le total du montant n'est pas correcte. "
                                + "Le total écrit est " + line.amount() + " alors que le calcul donne " + totalBuildingAmount + "\n";
                    }
                    if (totalBuildingDeduction != Double.parseDouble(line.deduction())) {
                        if (message.isEmpty()) {
                            message = "Pour l'immeuble :\n";
                        }
                        message += "Le total de la déduction n'est pas correcte. Le total écrit est " + line.deduction()
                                + " alors que le calcul donne " + totalBuildingDeduction + "\n";
                    }
                    if (totalBuildingRecovery != Double.parseDouble(line.recovery())) {
                        if (message.isEmpty()) {
                            message = "Pour l'immeuble :\n";
                        }
                        message += "Le total de la récuperation n'est pas correcte. Le total écrit est " + line.recovery()
                                + " alors que le calcul donne " + totalBuildingRecovery + "\n";
                    }
                    if (!message.isEmpty()) {
                        listOfMessage.add(message);
                    }
                }
            }
        }
        return listOfMessage;
    }

    public TreeMap<String, String> checkDocument(LineOfExpense[] listOfExpense, String pathOfDocuments) {
        TreeMap<String, String> listOfDocuments = new TreeMap<>();
        for (LineOfExpense lineOfExpense : listOfExpense) {
            if (lineOfExpense instanceof LineOfExpenseValue) {
                String document = ((LineOfExpenseValue) lineOfExpense).document();
                File pathDirectoryInvoice = new File(pathOfDocuments);
                File fileFound = null;
                File[] files = pathDirectoryInvoice.listFiles();
                if (null != files) {
                    for (File fichier : files) {
                        if (fichier.getName().startsWith(document)) {
                            fileFound = fichier;
                            break;
                        } else if (fichier.isDirectory()) {
                            fileFound = outilChek.getFileInDirectory(document, fichier);
                            if (fileFound != null) {
                                break;
                            }
                        }
                    }
                    String message;
                    if (fileFound == null) {
                        message = IMPOSSIBLE_DE_TROUVER_LA_PIECE + document + " pour le libellé : " + ((LineOfExpenseValue) lineOfExpense).label();

                    } else {
                        message = fileFound.getAbsoluteFile().toString();
                    }
                    listOfDocuments.put(document, message);
                }
            }
        }
        return listOfDocuments;
    }

    public void writeFileListOfExpense(LineOfExpense[] listOfExpense, TreeMap<String, String> listOfDocuments) {
        WriteFile writeFile = new WriteFile(PATH);
        writeFile.writeFileExcelListeDesDepenses(listOfExpense, listOfDocuments);
    }
}