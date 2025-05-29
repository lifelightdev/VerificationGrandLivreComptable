package life.light.write;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import static life.light.Constant.DATE_FORMATTER;
import static life.light.Constant.REPORT_DE;
import static life.light.write.WriteOutil.*;

public class WriteLine {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final String LE_GRAND_LIVRE = "le grand livre";
    public static final String LE_GRAND_LIVRE_DE_L_ETAT_DE_RAPPROCHEMENT = "le grand livre de l'état de rapprochement";
    public static final String LE_RELEVE_DE_COMPTE_BANQUE = "le relevé de compte banque";

    WriteCellStyle writeCellStyle = new WriteCellStyle();
    WriteCell writeCell = new WriteCell();
    WriteOutil writeOutil = new WriteOutil();

    public void getLineGrandLivre(Line lineOfLedger, Row row, boolean verif, String pathDirectoryInvoice) {
        boolean isWhite = row.getRowNum() % 2 == 0;
        CellStyle cellStyle = writeCellStyle.getCellStyle(row.getSheet().getWorkbook(), isWhite);
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleAmount(row.getSheet().getWorkbook(), isWhite);

        writeCell.addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle,
                lineOfLedger.toString(), "le numéro de compte", LE_GRAND_LIVRE);
        writeCell.addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle,
                lineOfLedger.toString(), "le libellé du compte", LE_GRAND_LIVRE);

        boolean isReportLine = lineOfLedger.label().contains(REPORT_DE);

        if (isReportLine) {
            writeCell.addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_DATE_OF_LEDGER, row, cellStyle);
        } else {
            writeCell.addCell(row, ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle, lineOfLedger.toString(),
                    "la piéce", LE_GRAND_LIVRE);
        }
        writeCell.addCell(row, ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle, lineOfLedger.toString(),
                "la date", LE_GRAND_LIVRE);

        if (isReportLine) {
            writeCell.addCellEmpty(ID_JOURNAL_OF_LEDGER, ID_CHECK_OF_LEDGER, row, cellStyle);
        } else {
            writeCell.addCell(row, ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle, lineOfLedger.toString(),
                    null, LE_GRAND_LIVRE);
            writeCell.addCell(row, ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(),
                    cellStyle, lineOfLedger.toString(), "le numéro de compte de la contrepartie",
                    LE_GRAND_LIVRE);
            writeCell.addCell(row, ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                    lineOfLedger.toString(), "le libellé de la contrepartie", LE_GRAND_LIVRE);
        }
        writeCell.addCell(row, ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle, lineOfLedger.toString(),
                null, LE_GRAND_LIVRE);
        writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle, lineOfLedger.toString(),
                "le libellé", LE_GRAND_LIVRE);

        Cell debitCell = writeCell.addDebitCell(lineOfLedger, row, cellStyleAmount);
        Cell creditCell = writeCell.addCreditCell(lineOfLedger, row, cellStyleAmount);
        writeCell.addSoldeCell(row, debitCell, creditCell, cellStyleAmount,
                ID_BALANCE_OF_LEDGER, lineOfLedger.label().contains(REPORT_DE), false);

        if (verif) {
            writeOutil.addVerifCells(lineOfLedger, row, cellStyle, pathDirectoryInvoice);
        }
    }

    public void getLineEtatRapprochement(Line lineOfLedger, Row row, BankLine bankLine, String message) {
        boolean isWhite = row.getRowNum() % 2 == 0;
        CellStyle cellStyle = writeCellStyle.getCellStyle(row.getSheet().getWorkbook(), isWhite);
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleAmount(row.getSheet().getWorkbook(), isWhite);

        if (lineOfLedger != null) {
            List<CellValues> values = getCellValuesOfLedgerInLineEtatRapprochement(lineOfLedger, cellStyle);
            writeCell.addCells(row, values, LE_GRAND_LIVRE_DE_L_ETAT_DE_RAPPROCHEMENT);
            writeCell.addDebitCell(lineOfLedger, row, cellStyleAmount);
            writeCell.addCreditCell(lineOfLedger, row, cellStyleAmount);
        }

        if (bankLine != null) {
            List<CellValues> values = getCellValuesObBankInEtatRapprochement(bankLine, cellStyle, cellStyleAmount);
            writeCell.addCells(row, values, LE_RELEVE_DE_COMPTE_BANQUE);
        }

        writeCell.addCell(row, ID_COMMENT_OF_RECONCILIATION, message, cellStyle, null, null, null);
    }

    private List<CellValues> getCellValuesObBankInEtatRapprochement(BankLine bankLine, CellStyle cellStyle, CellStyle cellStyleAmount) {
        List<CellValues> values = new ArrayList<>();
        values.add(new CellValues(ID_MONTH_OF_SATEMENT_OF_RECONCILIATION, bankLine.year() + "-" + bankLine.mounth(), cellStyle,
                bankLine.toString(), "le mois et l'année"));
        values.add(new CellValues(ID_ACOUNT_NUMBER_OF_RECONCILIATION, bankLine.account().account(), cellStyle,
                bankLine.toString(), "le numéro du compte"));
        values.add(new CellValues(ID_ACOUNT_LABEL_OF_RECONCILIATION, bankLine.account().label(), cellStyle,
                bankLine.toString(), "le libellé du compte"));
        values.add(new CellValues(ID_OPERATION_DATE_OF_RECONCILIATION, bankLine.operationDate().format(DATE_FORMATTER), cellStyle,
                bankLine.toString(), "la date de l'opération"));
        values.add(new CellValues(ID_VALUE_DATE_OF_RECONCILIATION, bankLine.valueDate().format(DATE_FORMATTER), cellStyle,
                bankLine.toString(), "la date de valeur"));
        values.add(new CellValues(ID_LABEL_OF_RECONCILIATION, bankLine.label(), cellStyle,
                bankLine.toString(), "le libellé"));
        values.add(new CellValues(ID_DEBIT_OF_RECONCILIATION, bankLine.debit().toString(), cellStyleAmount,
                bankLine.toString(), "le débit"));
        values.add(new CellValues(ID_CREDIT_OF_RECONCILIATION, bankLine.credit().toString(), cellStyleAmount,
                bankLine.toString(), "le crédit"));
        return values;
    }

    private static List<CellValues> getCellValuesOfLedgerInLineEtatRapprochement(Line lineOfLedger, CellStyle cellStyle) {
        List<CellValues> values = new ArrayList<>();
        values.add(new CellValues(ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle,
                lineOfLedger.toString(), "le numéro de compte"));
        values.add(new CellValues(ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle,
                lineOfLedger.toString(), "le libellé du compte"));
        values.add(new CellValues(ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle,
                lineOfLedger.toString(), "la piéce"));
        values.add(new CellValues(ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle,
                lineOfLedger.toString(), "la date"));
        values.add(new CellValues(ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle,
                lineOfLedger.toString(), null));
        values.add(new CellValues(ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(), cellStyle,
                lineOfLedger.toString(), "le numéro de compte de la contrepartie"));
        values.add(new CellValues(ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                lineOfLedger.toString(), "le libellé de la contrepartie"));
        values.add(new CellValues(ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle,
                lineOfLedger.toString(), null));
        values.add(new CellValues(ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle,
                lineOfLedger.toString(), "le libellé"));
        return values;
    }

    public void getTotalBuilding(TotalBuilding lineOfTotalBuildingInLedger, Row row, List<Integer> lineTotals) {

        CellStyle styleTotal = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle styleTotalAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        writeCell.addCellEmpty(ID_ACOUNT_NUMBER_OF_LEDGER, ID_LABEL_OF_LEDGER, row, styleTotal);
        writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfTotalBuildingInLedger.label(), styleTotal,
                lineOfTotalBuildingInLedger.toString(), "le libellé", "le grand livre");

        Cell debitCell = writeCell.addCellAmountOfTotalBuildingInLedger(row, ID_DEBIT_OF_LEDGER, lineTotals, styleTotalAmount);
        Cell creditCell = writeCell.addCellAmountOfTotalBuildingInLedger(row, ID_CREDIT_OF_LEDGER, lineTotals, styleTotalAmount);

        writeCell.addSoldeCell(row, debitCell, creditCell, styleTotalAmount, ID_BALANCE_OF_LEDGER, false, true);
    }

    public void getTotalAccount(TotalAccount lineOfTotalAccountInLedger, Row row, int lastRowNumTotal) {

        CellStyle cellStyle = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        List<CellValues> values = getCellValuesOfTotalAccount(lineOfTotalAccountInLedger, cellStyle);
        writeCell.addCells(row, values, LE_GRAND_LIVRE);

        writeCell.addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_LABEL_OF_LEDGER, row, cellStyle);

        Cell debitCell = writeCell.addCellDebitOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell creditCell = writeCell.addCellCreditOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell soldeCell = writeCell.addSoldeCell(row, debitCell, creditCell, cellStyleAmount, ID_BALANCE_OF_LEDGER, false, true);

        Cell cellVerif = row.createCell(ID_VERIFFICATION_OF_LEDGER);
        String amount = getAmountInLabelOfTotalLine(lineOfTotalAccountInLedger);
        String verif;
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        double solde = getRound(evaluator, soldeCell);
        double debitExcel = getRound(evaluator, debitCell);
        double debitGrandLivre = Double.parseDouble(lineOfTotalAccountInLedger.debit());
        double creditExcel = getRound(evaluator, creditCell);
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

        addCellMessageInTotalaccount(row, soldeCell, amount, creditCell, creditGrandLivre, debitCell, debitGrandLivre, cellStyle);
    }

    private double getRound(FormulaEvaluator evaluator, Cell soldeCell) {
        return (double) Math.round(evaluator.evaluate(soldeCell).getNumberValue() * 100) / 100;
    }

    private static String getAmountInLabelOfTotalLine(TotalAccount lineOfTotalAccountInLedger) {
        return lineOfTotalAccountInLedger.label()
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
    }

    private static void addCellMessageInTotalaccount(Row row, Cell soldeCell, String amount, Cell creditCell, double creditGrandLivre, Cell debitCell, double debitGrandLivre, CellStyle cellStyle) {
        Cell cellMessqage = row.createCell(ID_COMMENT_OF_LEDGER);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount) + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        //cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(cellStyle);
    }

    private static List<CellValues> getCellValuesOfTotalAccount(TotalAccount lineOfTotalAccountInLedger, CellStyle cellStyle) {
        List<CellValues> values = new ArrayList<>();
        values.add(new CellValues(ID_ACOUNT_NUMBER_OF_LEDGER, lineOfTotalAccountInLedger.account().account(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le numéro de compte"));
        values.add(new CellValues(ID_ACOUNT_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.account().label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé du compte"));
        values.add(new CellValues(ID_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé"));
        return values;
    }

    public void getCellsEntete(Sheet sheet, String[] header) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = writeCellStyle.getCellStyleHeader(sheet.getWorkbook().createCellStyle());
        for (String label : header) {
            writeCell.addCell(headerRow, index++, label, styleHeader, "", null, null);
        }
    }

    public void getCellsEnteteEtatRapprochement(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        for (String label : NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT) {
            writeCell.addCell(headerRow, index++, label,
                    writeCellStyle.getCellStyleHeader(sheet.getWorkbook().createCellStyle()), "",
                    null, null);
        }
    }

    private void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row, int lastRowNumTotal, List<Integer> listId) {
        Cell cell;
        String message = "";
        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
        FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        if (!line.key().isEmpty()) {
            cell = row.createCell(ID_LABEL_OF_LIST_OF_EXPENSES);
            if (line.type().equals(TypeOfExpense.Key)) {
                cell.setCellValue("Total de la clé : " + line.key());
                cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
                cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
                StringBuilder sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                double cellvalue = getRound(evaluator, cell);
                double amount = Double.parseDouble(line.amount());
                row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
                if (!Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "Le montant est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier.";
                }
                cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
                sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                cellvalue = getRound(evaluator, cell);
                amount = Double.parseDouble(line.deduction());
                if (!Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "La déduction est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
                }
                cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
                sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                cellvalue = getRound(evaluator, cell);
                amount = Double.parseDouble(line.recovery());
                if (!Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "La récupération est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
                }
            }
            if (line.type().equals(TypeOfExpense.Nature)) {
                cell.setCellValue("Total de la nature : " + line.key());
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                if (!line.amount().isEmpty()) {
                    message += getAddAmount(ID_AMOUNT_OF_LIST_OF_EXPENSES, row, lastRowNumTotal, evaluator, line.amount(), "Le montant");
                }
                if (!line.deduction().isEmpty()) {
                    message += getAddAmount(ID_DEDUCTION_OF_LIST_OF_EXPENSES, row, lastRowNumTotal, evaluator, line.deduction(), "La déduction");
                }
                if (!line.recovery().isEmpty()) {
                    message += getAddAmount(ID_RECOVERY_OF_LIST_OF_EXPENSES, row, lastRowNumTotal, evaluator, line.recovery(), "La récupération");
                }
            }
            if (line.type().equals(TypeOfExpense.Building)) {
                cell.setCellValue("Total de l'immeuble : " + line.key());
                cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
                cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
                StringBuilder sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                double cellvalue = getRound(evaluator, cell);
                double amount = Double.parseDouble(line.amount());
                if (Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "Le montant est de " + Double.parseDouble(line.amount()) + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
                }
                cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
                sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                cellvalue = getRound(evaluator, cell);
                amount = Double.parseDouble(line.deduction());
                if (Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "La déduction est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
                }
                cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
                sum = new StringBuilder();
                for (Integer numRow : listId) {
                    sum.append(CellReference.convertNumToColString(cell.getColumnIndex())).append(numRow).append("+");
                }
                cell.setCellFormula(sum.substring(0, sum.lastIndexOf("+")));
                cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
                cellvalue = getRound(evaluator, cell);
                amount = Double.parseDouble(line.recovery());
                if (Double.toString(amount).equals(Double.toString(cellvalue))) {
                    cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
                    message += "La récupération est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
                }
            }
        }
        cell = row.createCell(ID_COMMENT_OF_LIST_OF_EXPENSES);
        cell.setCellValue(message);
        cell.setCellStyle(writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook()));
    }

    private String getAddAmount(int Id_colum, Row row, int lastRowNumTotal, FormulaEvaluator evaluator, String value, String name) {
        Cell cell;
        double amount = Double.parseDouble(value);
        cell = row.createCell(Id_colum);
        CellAddress cellAddressFirst = new CellAddress(lastRowNumTotal - 1, cell.getAddress().getColumn());
        CellAddress cellAddressEnd = new CellAddress(cell.getAddress().getRow() - 1, cell.getAddress().getColumn());
        cell.setCellFormula("SUM(" + cellAddressFirst + ":" + cellAddressEnd + ")");
        cell.setCellStyle(writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        double cellvalue = getRound(evaluator, cell);
        String message = "";
        if (!Double.toString(amount).equals(Double.toString(cellvalue))) {
            cell.setCellStyle(writeCellStyle.getCellStyleVerifRed(row.getSheet().getWorkbook().createCellStyle()));
            message = name + " est de " + amount + " dans le PDF au lieu de " + cellvalue + " dans ce fichier. ";
        }
        return message;
    }

    public void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row, int lastRowNumTotalNature) {
        if (line.type().equals(TypeOfExpense.Nature)) {
            if (lastRowNumTotalNature > 0) {
                getLineOfExpenseTotal(line, row, lastRowNumTotalNature, null);
            } else {
                LOGGER.error("Il manque l'ID de la ligne du total de la nature précédente.");
            }
        }
    }

    public void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row, List<Integer> listIdLineTotal) {
        if (line.type().equals(TypeOfExpense.Key)) {
            if (!listIdLineTotal.isEmpty()) {
                getLineOfExpenseTotal(line, row, 0, listIdLineTotal);
            } else {
                LOGGER.error("Il manque la liste des ID des lignes total Nature.");
            }
        }
        if (line.type().equals(TypeOfExpense.Building)) {
            if (!listIdLineTotal.isEmpty()) {
                getLineOfExpenseTotal(line, row, 0, listIdLineTotal);
            } else {
                LOGGER.error("Il manque la liste des ID des lignes total clé.");
            }
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

    public void getLineOfExpense(LineOfExpense line, Row row, String pathDirectoryInvoice) {
        CreationHelper createHelper = row.getSheet().getWorkbook().getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);

        boolean isWhite = row.getRowNum() % 2 == 0;
        CellStyle cellStyle = writeCellStyle.getCellStyle(row.getSheet().getWorkbook(), isWhite);
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleAmount(row.getSheet().getWorkbook(), isWhite);

        Cell cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        if (writeOutil.isDouble(line.document())) {
            cell.setCellValue(Double.parseDouble(line.document()));
        } else {
            cell.setCellValue(line.document());
        }
        cell.setCellStyle(cellStyle);

        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellValue(line.date().format(DATE_FORMATTER));
        cell.setCellStyle(cellStyle);

        cell = row.createCell(ID_LABEL_OF_LIST_OF_EXPENSES);
        cell.setCellValue(line.label());
        cell.setCellStyle(cellStyle);

        cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
        cell.setCellValue(Double.parseDouble(line.amount()));
        cell.setCellStyle(cellStyleAmount);

        if (!line.deduction().isEmpty()) {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.deduction()));
            cell.setCellStyle(cellStyleAmount);
        } else {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellStyle(cellStyleAmount);
        }

        if (!line.recovery().isEmpty()) {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.recovery()));
            cell.setCellStyle(cellStyleAmount);
        } else {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellStyle(cellStyleAmount);
        }

        cell = row.createCell(ID_COMMENT_OF_LIST_OF_EXPENSES);

        String message = writeOutil.getMessageFindDocument(line.document(), link, pathDirectoryInvoice);
        if (message.startsWith(IMPOSSIBLE_DE_TROUVER_LA_PIECE)) {
            message = message + " avec ce libellé " + line.label();
        }
        cell.setCellValue(message);
        cell.setCellStyle(cellStyle);
    }

    public void getListOfDocumentMissing(TreeMap<String, String> ligneOfDocumentMissing, Sheet sheetDocument) {
        int index = 0;
        for (Map.Entry<String, String> entry : ligneOfDocumentMissing.entrySet()) {
            Row row = sheetDocument.createRow(index++);
            Cell cellD = row.createCell(0);
            cellD.setCellValue(Integer.parseInt(entry.getKey()));
            Cell cellM = row.createCell(1);
            cellM.setCellValue(entry.getValue());
        }
    }
}
