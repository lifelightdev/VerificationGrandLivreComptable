package life.light.write;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class WriteOutil {

    private WriteCellStyle writeCellStyle = new WriteCellStyle();

    private static final Logger LOGGER = LogManager.getLogger();

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    public static final String[] NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES = {"Pièce", "Date", "Libellé", "Montant",
            "Déduction", "Récuperation", "Commentaire"};
    public static final int ID_DOCUMENT_OF_LIST_OF_EXPENSES = 0;
    private static final int ID_DATE_OF_LIST_OF_EXPENSES = 1;
    private static final int ID_LABEL_OF_LIST_OF_EXPENSES = 2;
    private static final int ID_AMOUNT_OF_LIST_OF_EXPENSES = 3;
    private static final int ID_DEDUCTION_OF_LIST_OF_EXPENSES = 4;
    private static final int ID_RECOVERY_OF_LIST_OF_EXPENSES = 5;
    public static final int ID_COMMENT_OF_LIST_OF_EXPENSES = 6;

    protected static final String[] NOM_ENTETE_COLONNE_GRAND_LIVRE = {"Compte", "Intitulé du compte", "Pièce", "Date",
            "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "Solde (Calculé)", "Vérification", "Commentaire"};
    private static final int ID_ACOUNT_NUMBER_OF_LEDGER = 0;
    private static final int ID_ACOUNT_LABEL_OF_LEDGER = 1;
    public static final int ID_DOCUMENT_OF_LEDGER = 2;
    private static final int ID_DATE_OF_LEDGER = 3;
    private static final int ID_JOURNAL_OF_LEDGER = 4;
    private static final int ID_COUNTERPART_NUMBER_OF_LEDGER = 5;
    private static final int ID_COUNTERPART_LABEL_OF_LEDGER = 6;
    private static final int ID_CHECK_OF_LEDGER = 7;
    private static final int ID_LABEL_OF_LEDGER = 8;
    private static final int ID_DEBIT_OF_LEDGER = 9;
    private static final int ID_CREDIT_OF_LEDGER = 10;
    private static final int ID_BALANCE_OF_LEDGER = 11;
    private static final int ID_VERIFFICATION_OF_LEDGER = 12;
    public static final int ID_COMMENT_OF_LEDGER = 13;

    protected static final String[] NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT = {"Compte", "Intitulé du compte", "Pièce",
            "Date", "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "----",
            "Mois du relevé", "Compte", "Intitulé du compte", "Date de l'opération", "Date de valeur", "Libellé",
            "Débit", "Crédit", "Commentaire"};
    private static final int ID_MONTH_OF_SATEMENT_OF_RECONCILIATION = 12;
    private static final int ID_ACOUNT_NUMBER_OF_RECONCILIATION = 13;
    private static final int ID_ACOUNT_LABEL_OF_RECONCILIATION = 14;
    private static final int ID_OPERATION_DATE_OF_RECONCILIATION = 15;
    private static final int ID_VALUE_DATE_OF_RECONCILIATION = 16;
    private static final int ID_LABEL_OF_RECONCILIATION = 17;
    private static final int ID_DEBIT_OF_RECONCILIATION = 18;
    private static final int ID_CREDIT_OF_RECONCILIATION = 19;
    private static final int ID_COMMENT_OF_RECONCILIATION = 20;

    private static final String REPORT_DE = "Report de";
    public static final String IMPOSSIBLE_DE_TROUVER_LA_PIECE = "Impossible de trouver la pièce ";

    private static final String CLASSE_6 = "6";
    public static final String KO = "KO";
    public static final String OK = "OK";

    public void writeWorkbook(String fileName, Workbook workbook) {
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
            LOGGER.info("L'écriture du fichier {} est terminée.", fileName);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public void getTotalBuilding(TotalBuilding lineOfTotalBuildingInLedger, Row row, List<Integer> lineTotals) {

        CellStyle styleTotal = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle styleTotalAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        addCellEmpty(ID_ACOUNT_NUMBER_OF_LEDGER, ID_LABEL_OF_LEDGER, row, styleTotal);
        addCell(row, ID_LABEL_OF_LEDGER, lineOfTotalBuildingInLedger.label(), styleTotal,
                lineOfTotalBuildingInLedger.toString(), "le libellé", "le grand livre");

        Cell debitCell = addCellAmountOfTotalBuildingInLedger(row, ID_DEBIT_OF_LEDGER, lineTotals, styleTotalAmount);
        Cell creditCell = addCellAmountOfTotalBuildingInLedger(row, ID_CREDIT_OF_LEDGER, lineTotals, styleTotalAmount);

        addCellBalanceOfTotalInLedger(row, debitCell, creditCell, styleTotalAmount);
    }

    private Cell addCellBalanceOfTotalInLedger(Row row, Cell debitCell, Cell creditCell, CellStyle styleTotalAmount) {
        Cell soldeCell = row.createCell(ID_BALANCE_OF_LEDGER);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(styleTotalAmount);
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        return soldeCell;
    }

    private Cell addCellAmountOfTotalBuildingInLedger(Row row, int idDebitOfLedger, List<Integer> lineTotals, CellStyle styleTotalAmount) {
        Cell debitCell = row.createCell(idDebitOfLedger);
        StringBuilder sumDebit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumDebit.append(CellReference.convertNumToColString(debitCell.getColumnIndex())).append(numRow).append("+");
        }
        debitCell.setCellFormula(sumDebit.substring(0, sumDebit.lastIndexOf("+")));
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(styleTotalAmount);
        return debitCell;
    }

    private void addCell(Row row, int idColum, String value, CellStyle style, String line, String name, String place) {
        Cell cell = row.createCell(idColum);
        if (isDouble(value)) {
            cell.setCellValue(Double.parseDouble(value));
        } else if (!value.isEmpty()) {
            cell.setCellValue(value);
        } else {
            if (name != null && !name.isEmpty()) {
                LOGGER.error("Il manque {} dans {} à la ligne {}", name, place, line);
            }
        }
        cell.setCellStyle(style);
    }

    private void addCellEmpty(int idFirstColum, int idLastColum, Row row, CellStyle style) {
        for (int idCell = idFirstColum; idCell < idLastColum; idCell++) {
            Cell cell = row.createCell(idCell);
            cell.setCellStyle(style);
        }
    }

    public void getTotalAccount(TotalAccount lineOfTotalAccountInLedger, Row row, int lastRowNumTotal) {

        CellStyle cellStyle = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfTotalAccountInLedger.account().account(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le numéro de compte", "le grand livre");
        addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.account().label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé du compte", "le grand livre");
        addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_LABEL_OF_LEDGER, row, cellStyle);
        addCell(row, ID_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé", "le grand livre");

        Cell debitCell = addCellDebitOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell creditCell = addCellCreditOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell soldeCell = addCellBalanceOfTotalInLedger(row, debitCell, creditCell, cellStyleAmount);

        Cell cellVerif = row.createCell(ID_VERIFFICATION_OF_LEDGER);
        String amount = lineOfTotalAccountInLedger.label()
                .replace("Total", " ")
                .replace("compte", " ")
                .replace("(Solde", " ")
                .replace("créditeur", " ")
                .replace("débiteur", " ")
                .replace(" : ", "")
                .replace(":", "")
                .replace("€", "")
                .replace(")", "")
                .replace(" ", "")
                .trim();
        String verif;
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        double solde = (double) Math.round(evaluator.evaluate(soldeCell).getNumberValue() * 100) / 100;
        double debitExcel = (double) Math.round(evaluator.evaluate(debitCell).getNumberValue() * 100) / 100;
        double debitGrandLivre = Double.parseDouble(lineOfTotalAccountInLedger.debit());
        double creditExcel = (double) Math.round(evaluator.evaluate(creditCell).getNumberValue() * 100) / 100;
        double creditGrandLivre = Double.parseDouble(lineOfTotalAccountInLedger.credit());
        if (Double.parseDouble(amount) == solde) {
            if (debitExcel == debitGrandLivre) {
                if (creditExcel == creditGrandLivre) {
                    verif = "OK";
                    cellVerif.setCellStyle(cellStyleAmount);
                } else {
                    verif = KO;
                }
            } else {
                verif = KO;
            }
        } else {
            verif = KO;
        }
        if (verif.equals(KO)) {
            cellVerif.setCellStyle(writeCellStyle.getCellStyleVerifRed(cellStyle));
            LOGGER.info("Vérification du total du compte [{}] Solde PDF [{}] Solde Excel [{}] débit PDF [{}] débit Excel [{}] crédit PDF [{}] crédit Excel [{}] ligne [{}]",
                    lineOfTotalAccountInLedger.account().account(), amount, solde,
                    debitGrandLivre, debitExcel,
                    creditGrandLivre, creditExcel,
                    lineOfTotalAccountInLedger);
        }
        cellVerif.setCellValue(verif);

        Cell cellMessqage = row.createCell(ID_COMMENT_OF_LEDGER);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount) + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        //cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(cellStyle);
    }

    private Cell addCellCreditOfTotalAccountInLedger(Row row, int lastRowNumTotal, CellStyle cellStyleAmount) {
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        creditCell.setCellStyle(cellStyleAmount);
        CellAddress creditCellAddressFirst = new CellAddress(lastRowNumTotal + 1, creditCell.getAddress().getColumn());
        CellAddress creditCellAddressEnd = new CellAddress(creditCell.getAddress().getRow() - 1, creditCell.getAddress().getColumn());
        creditCell.setCellFormula("SUM(" + creditCellAddressFirst + ":" + creditCellAddressEnd + ")");
        return creditCell;
    }

    private Cell addCellDebitOfTotalAccountInLedger(Row row, int lastRowNumTotal, CellStyle cellStyleAmount) {
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        CellAddress debitCellAddressFirst = new CellAddress(lastRowNumTotal + 1, debitCell.getAddress().getColumn());
        CellAddress debitCellAddressEnd = new CellAddress(debitCell.getAddress().getRow() - 1, debitCell.getAddress().getColumn());
        debitCell.setCellFormula("SUM(" + debitCellAddressFirst + ":" + debitCellAddressEnd + ")");
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(cellStyleAmount);
        return debitCell;
    }

    public void getLineGrandLivre(Line lineOfLedger, Row row, boolean verif, String pathDirectoryInvoice) {
        CellStyle cellStyle;
        CellStyle cellStyleAmount;
        if (row.getRowNum() % 2 == 0) {
            cellStyle = writeCellStyle.getCellStyleWhite(row.getSheet().getWorkbook());
            cellStyleAmount = writeCellStyle.getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            cellStyle = writeCellStyle.getCellStyleBlue(row.getSheet().getWorkbook());
            cellStyleAmount = writeCellStyle.getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }

        addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle, lineOfLedger.toString(),
                "le numéro de compte", "legrand livre");
        addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle, lineOfLedger.toString(),
                "le libellé du compte", "le grand livre");
        if (!lineOfLedger.label().contains(REPORT_DE)) {
            addCell(row, ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle, lineOfLedger.toString(),
                    "la piéce", "le grand livre");
        } else {
            addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_DATE_OF_LEDGER, row, cellStyle);
        }
        addCell(row, ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle, lineOfLedger.toString(), "la date",
                "le grand livre");

        if (!lineOfLedger.label().contains(REPORT_DE)) {
            addCell(row, ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle, lineOfLedger.toString(), null,
                    "le grand livre");
            addCell(row, ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte de la contrepartie", "le grand livre");
            addCell(row, ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                    lineOfLedger.toString(), "le libellé de la contrepartie", "le grand livre");
        } else {
            addCellEmpty(ID_JOURNAL_OF_LEDGER, ID_CHECK_OF_LEDGER, row, cellStyle);
        }
        addCell(row, ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle, lineOfLedger.toString(), null,
                "le grand livre");
        addCell(row, ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle, lineOfLedger.toString(),
                "le libellé", "le grand livre");

        Cell debitCell = addDebitCell(lineOfLedger, row, cellStyleAmount);
        Cell creditCell = addCreditCell(lineOfLedger, row, cellStyleAmount);
        addSoldeCell(lineOfLedger, row, cellStyleAmount, creditCell, debitCell);

        if (verif) {
            addVerifCells(lineOfLedger, row, cellStyle, pathDirectoryInvoice);
        }
    }

    public void getLineEtatRapprochement(Line lineOfLedger, Row row, BankLine bankLine, String message) {
        CellStyle cellStyle;
        CellStyle cellStyleAmount;
        if (row.getRowNum() % 2 == 0) {
            cellStyle = writeCellStyle.getCellStyleWhite(row.getSheet().getWorkbook());
            cellStyleAmount = writeCellStyle.getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            cellStyle = writeCellStyle.getCellStyleBlue(row.getSheet().getWorkbook());
            cellStyleAmount = writeCellStyle.getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }
        if (lineOfLedger != null) {
            String place = "le grand livre de l'état de rapprochement";
            addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte", place);
            addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle, lineOfLedger.toString(),
                    "le libellé de compte", place);
            addCell(row, ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle, lineOfLedger.toString(),
                    "la piéce", place);
            addCell(row, ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle, lineOfLedger.toString(),
                    "la date", place);
            addCell(row, ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle, lineOfLedger.toString(), null,
                    place);
            addCell(row, ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte de la contrepartie", place);
            addCell(row, ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                    lineOfLedger.toString(), "le libellé de la contrepartie", place);
            addCell(row, ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle, lineOfLedger.toString(),
                    null, place);
            addCell(row, ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle, lineOfLedger.toString(),
                    "le libellé", place);
            addDebitCell(lineOfLedger, row, cellStyleAmount);
            addCreditCell(lineOfLedger, row, cellStyleAmount);
        }
        if (bankLine != null) {
            String place = "le relevé de compte banque";
            addCell(row, ID_MONTH_OF_SATEMENT_OF_RECONCILIATION, bankLine.year() + "-" + bankLine.mounth(),
                    cellStyle, bankLine.toString(), "le mois et l'année", place);
            addCell(row, ID_ACOUNT_NUMBER_OF_RECONCILIATION, bankLine.account().account(), cellStyle,
                    bankLine.toString(), "le numéro du compte", place);
            addCell(row, ID_ACOUNT_LABEL_OF_RECONCILIATION, bankLine.account().label(), cellStyle, bankLine.toString(),
                    "le libellé du compte", place);
            addCell(row, ID_OPERATION_DATE_OF_RECONCILIATION, bankLine.operationDate().format(DATE_FORMATTER),
                    cellStyle, bankLine.toString(), "la date de l'opération", place);
            addCell(row, ID_VALUE_DATE_OF_RECONCILIATION, bankLine.valueDate().format(DATE_FORMATTER), cellStyle,
                    bankLine.toString(), "la date de valeur", place);
            addCell(row, ID_LABEL_OF_RECONCILIATION, bankLine.label(), cellStyle, bankLine.toString(),
                    "le libellé", place);

            Cell debitReleveCell = row.createCell(ID_DEBIT_OF_RECONCILIATION);
            debitReleveCell.setCellValue(bankLine.debit());
            debitReleveCell.setCellStyle(cellStyleAmount);

            Cell creditReleveCell = row.createCell(ID_CREDIT_OF_RECONCILIATION);
            creditReleveCell.setCellValue(bankLine.credit());
            creditReleveCell.setCellStyle(cellStyleAmount);
        }
        addCell(row, ID_COMMENT_OF_RECONCILIATION, message, cellStyle, null, null, null);
    }

    private void addVerifCells(Line grandLivre, Row row, CellStyle style, String pathDirectoryInvoice) {
        String message = "";
        CreationHelper createHelper = row.getSheet().getWorkbook().getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
        Cell verifCell = row.createCell(ID_VERIFFICATION_OF_LEDGER);
        if (grandLivre.label().startsWith(REPORT_DE)) {
            message = getMessageVerifLineReport(grandLivre, verifCell, style, message);
        } else {
            if (grandLivre.account().account().startsWith(CLASSE_6)
                    || (grandLivre.account().account().startsWith("401")) && (grandLivre.accountCounterpart().account().startsWith(CLASSE_6))) {
                message = getMessageFindDocument(grandLivre.document(), link, pathDirectoryInvoice);
                if (message.startsWith("Impossible")) {
                    message = message + " pour le compte " + grandLivre.account().account() + " avec ce Libellé d'opération " + grandLivre.label();
                }
            }
            if (grandLivre.accountCounterpart() == null) {
                verifCell.setCellValue(KO);
                message = "Il n'y a pas de Contrepartie";
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre);
            }
            if (grandLivre.debit().isEmpty() && grandLivre.credit().isEmpty()) {
                verifCell.setCellValue(KO);
                message = "Il n'y a aucun montant";
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre.toString());
            }
        }
        if (verifCell.getStringCellValue().equals(KO)) {
            verifCell.setCellStyle(writeCellStyle.getCellStyleVerifRed(style));
        } else {
            verifCell.setCellStyle(style);
        }

        Cell messageCell = row.createCell(ID_COMMENT_OF_LEDGER);
        if (link.getAddress() != null) {
            messageCell.setHyperlink(link);
        }
        messageCell.setCellValue(message.trim());
        messageCell.setCellStyle(style);
    }

    private String getMessageFindDocument(String document, Hyperlink link, String thePathDirectoryInvoice) {
        String message = "";
        File pathDirectoryInvoice = new File(thePathDirectoryInvoice);
        File fileFound = null;
        File[] files = pathDirectoryInvoice.listFiles();
        if (null != files) {
            for (File fichier : files) {
                if (fichier.getName().startsWith(document)) {
                    fileFound = fichier;
                    break;
                } else if (fichier.isDirectory()) {
                    fileFound = getFileInDirectory(document, fichier);
                    if (fileFound != null) {
                        break;
                    }
                }
            }
            if (fileFound != null) {
                message = fileFound.getAbsoluteFile().toString().replace("F:", "D:");
                link.setAddress(fileFound.toURI().toString().replace("F:", "D:"));
            } else {
                message = IMPOSSIBLE_DE_TROUVER_LA_PIECE + document
                        + " dans le dossier : " + pathDirectoryInvoice;
            }
        }
        return message;
    }

    private static File getFileInDirectory(String document, File file) {
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

    private String getMessageVerifLineReport(Line grandLivre, Cell verifCell, CellStyle style, String message) {
        verifCell.setCellStyle(style);
        String amount = getAmountInLineReport(grandLivre);
        if (isDouble(amount)) {
            double amountDouble = Double.parseDouble(amount);
            DecimalFormat df = new DecimalFormat("#.00");
            double debit = 0;
            if (isDouble(grandLivre.debit())) {
                debit = Double.parseDouble(grandLivre.debit());
            }
            double credit = 0;
            if (isDouble(grandLivre.credit())) {
                credit = Double.parseDouble(grandLivre.credit());
            }
            String nombreFormate = df.format((debit - credit)).replace(",", ".");
            if (amountDouble == Double.parseDouble(nombreFormate)) {
                verifCell.setCellValue("OK");
                verifCell.setCellStyle(style);
            } else {
                message = "Le montant du report est de %s le solde est de  %s le débit est de %s le credit est de %s"
                        .formatted(amount, Double.parseDouble(nombreFormate), Double.parseDouble(grandLivre.debit()), Double.parseDouble(grandLivre.credit()));
                LOGGER.info("{} sur le compte {}", message, grandLivre.account().account());
                verifCell.setCellValue(KO);
            }
        }
        return message;
    }

    private String getAmountInLineReport(Line grandLivre) {
        return grandLivre.label().substring(REPORT_DE.length(), grandLivre.label().length() - 1).trim().replace(" ", "");
    }

    private void addSoldeCell(Line grandLivre, Row row, CellStyle style, Cell creditCell, Cell debitCell) {
        Cell soldeCell = row.createCell(ID_BALANCE_OF_LEDGER);
        String formule;
        if (grandLivre.label().startsWith(REPORT_DE) || row.getRowNum() == 1) {
            formule = creditCell.getAddress() + "-" + debitCell.getAddress();
        } else {
            int rowIndex = soldeCell.getRowIndex() - 1;
            int col = soldeCell.getColumnIndex();
            CellAddress beforeSoldeCellAddress = new CellAddress(rowIndex, col);
            formule = beforeSoldeCellAddress + "+" + debitCell.getAddress() + "-" + creditCell.getAddress();
        }
        soldeCell.setCellFormula(formule);
        soldeCell.setCellStyle(style);
    }

    private Cell addCreditCell(Line grandLivre, Row row, CellStyle style) {
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        creditCell.setCellValue(grandLivre.credit());
        if (isDouble(grandLivre.credit())) {
            creditCell.setCellValue(Double.parseDouble(grandLivre.credit()));
        } else {
            creditCell.setCellValue(0);
        }
        creditCell.setCellStyle(style);
        return creditCell;
    }

    private Cell addDebitCell(Line grandLivre, Row row, CellStyle style) {
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        debitCell.setCellValue(grandLivre.debit());
        if (isDouble(grandLivre.debit())) {
            debitCell.setCellValue(Double.parseDouble(grandLivre.debit()));
        } else {
            debitCell.setCellValue(0D);
        }
        debitCell.setCellStyle(style);
        return debitCell;
    }

    public boolean isDouble(String str) {
        if (str == null || str.isEmpty()) {
            return false; // Une chaîne nulle ou vide ne peut pas être un double
        }
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException _) {
            return false;
        }
    }

    public void getCellsEnteteGrandLivre(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_GRAND_LIVRE) {
            addCell(headerRow, index++, label, styleHeader, "", null, null);
        }
    }

    public void getCellsEnteteListeDesDepenses(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle cellStyleHeader = writeCellStyle.getCellStyleEntete(sheet.getWorkbook().createCellStyle());
        for (String label : NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES) {
            addCell(headerRow, index++, label, cellStyleHeader, "", null, null);
        }
    }

    public void getCellsEnteteEtatRapprochement(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT) {
            addCell(headerRow, index++, label, writeCellStyle.getCellStyleEntete(styleHeader), "", null, null);
        }
    }

    public void getLineOfExpenseKey(LineOfExpenseKey line, Row row) {
        Cell cell;
        for (int index = 0; index < NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES.length; index++) {
            cell = row.createCell(index);
            if (index == ID_LABEL_OF_LIST_OF_EXPENSES) {
                cell.setCellValue(line.label() + " : " + line.key() + " " + line.value());
            }
            cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        }
    }

    public void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row) {
        Cell cell;
        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        if (!line.key().isEmpty()) {
            cell = row.createCell(ID_LABEL_OF_LIST_OF_EXPENSES);
            if (line.type().equals(TypeOfExpense.Key)) {
                cell.setCellValue("Total de la clé : " + line.key());
            }
            if (line.type().equals(TypeOfExpense.Nature)) {
                cell.setCellValue("Total de la nature : " + line.key());
            }
            if (line.type().equals(TypeOfExpense.Building)) {
                cell.setCellValue("Total de l'immeuble : " + line.key());
            }
            cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        }
        if (!line.amount().isEmpty()) {
            cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.amount()));
            cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        if (!line.deduction().isEmpty()) {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.deduction()));
            cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        if (!line.recovery().isEmpty()) {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.recovery()));
            cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        cell = row.createCell(ID_COMMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
    }

    public void getLineOfExpense(LineOfExpense line, Row row, String pathDirectoryInvoice) {
        CreationHelper createHelper = row.getSheet().getWorkbook().getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);

        CellStyle styleColor;
        CellStyle styleAmountColor;
        if (row.getRowNum() % 2 == 0) {
            styleColor = writeCellStyle.getCellStyleWhite(row.getSheet().getWorkbook());
            styleAmountColor = writeCellStyle.getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            styleColor = writeCellStyle.getCellStyleBlue(row.getSheet().getWorkbook());
            styleAmountColor = writeCellStyle.getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }

        Cell cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        if (isDouble(line.document())) {
            cell.setCellValue(Double.parseDouble(line.document()));
        } else {
            cell.setCellValue(line.document());
        }
        cell.setCellStyle(styleColor);

        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellValue(line.date().format(DATE_FORMATTER));
        cell.setCellStyle(styleColor);

        cell = row.createCell(ID_LABEL_OF_LIST_OF_EXPENSES);
        cell.setCellValue(line.label());
        cell.setCellStyle(styleColor);

        cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
        cell.setCellValue(Double.parseDouble(line.amount()));
        cell.setCellStyle(styleAmountColor);

        if (!line.deduction().isEmpty()) {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.deduction()));
            cell.setCellStyle(styleAmountColor);
        } else {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellStyle(styleAmountColor);
        }

        if (!line.recovery().isEmpty()) {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.recovery()));
            cell.setCellStyle(styleAmountColor);
        } else {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellStyle(styleAmountColor);
        }

        cell = row.createCell(ID_COMMENT_OF_LIST_OF_EXPENSES);
        String message = getMessageFindDocument(line.document(), link, pathDirectoryInvoice);
        if (message.startsWith(IMPOSSIBLE_DE_TROUVER_LA_PIECE)) {
            message = message + " avec ce libellé " + line.label();
        }
        cell.setCellValue(message);
        cell.setCellStyle(styleColor);
    }

    public void autoSizeCollum(int numberOfColumns, Sheet sheet) {
        for (int idCollum = 0; idCollum < numberOfColumns; idCollum++) {
            sheet.autoSizeColumn(idCollum);
        }
        sheet.createFreezePane(0, 1);
    }

    private void writeListOfDocumentMissing(TreeMap<String, String> ligneOfDocumentMissing, Sheet sheetDocument) {
        int index = 0;
        for (Map.Entry<String, String> entry : ligneOfDocumentMissing.entrySet()) {
            Row row = sheetDocument.createRow(index++);
            Cell cellD = row.createCell(0);
            cellD.setCellValue(Integer.parseInt(entry.getKey()));
            Cell cellM = row.createCell(1);
            cellM.setCellValue(entry.getValue());
        }
    }

    private TreeMap<String, String> getListOfDocumentMissing(Sheet sheet, int idCellComment, int idCellDocumment) {
        TreeMap<String, String> ligneOfDocumentMissing = new TreeMap<>();
        for (Row row : sheet) {
            boolean commentCellIsNotNull = row.getCell(idCellComment) != null;
            if (commentCellIsNotNull) {
                boolean commmentCellIsCellTypeString = row.getCell(idCellComment).getCellType() == CellType.STRING;
                if (commmentCellIsCellTypeString) {
                    boolean commentCellContainsDocumentMissing = row.getCell(idCellComment).getStringCellValue().contains(IMPOSSIBLE_DE_TROUVER_LA_PIECE);
                    if (commentCellContainsDocumentMissing) {
                        String document;
                        if (row.getCell(idCellDocumment).getCellType() == CellType.NUMERIC) {
                            document = String.valueOf(row.getCell(idCellDocumment).getNumericCellValue());
                        } else {
                            document = row.getCell(idCellDocumment).getStringCellValue();
                        }
                        String message = row.getCell(idCellComment).getStringCellValue();
                        ligneOfDocumentMissing.put(document.replace(".0", ""), message);
                    }
                }
            }
        }
        return ligneOfDocumentMissing;
    }

    public void writeDocumentMission(Workbook workbook, Sheet sheet, int idCellComment, int idCellDocumment) {
        Sheet sheetDocument = workbook.createSheet("Pieces manquante");
        TreeMap<String, String> ligneOfDocumentMissing = getListOfDocumentMissing(sheet, idCellComment, idCellDocumment);
        writeListOfDocumentMissing(ligneOfDocumentMissing, sheetDocument);
    }
}
