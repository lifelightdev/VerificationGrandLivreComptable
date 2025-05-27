package life.light.write;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

import static life.light.write.WriteOutil.*;

public class WriteLine {

    private static final Logger LOGGER = LogManager.getLogger();

    WriteCellStyle writeCellStyle = new WriteCellStyle();
    WriteCell writeCell = new WriteCell();
    WriteOutil writeOutil = new WriteOutil();

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

        writeCell.addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle, lineOfLedger.toString(),
                "le numéro de compte", "legrand livre");
        writeCell.addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle, lineOfLedger.toString(),
                "le libellé du compte", "le grand livre");
        if (!lineOfLedger.label().contains(REPORT_DE)) {
            writeCell.addCell(row, ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle, lineOfLedger.toString(),
                    "la piéce", "le grand livre");
        } else {
            writeCell.addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_DATE_OF_LEDGER, row, cellStyle);
        }
        writeCell.addCell(row, ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle, lineOfLedger.toString(), "la date",
                "le grand livre");

        if (!lineOfLedger.label().contains(REPORT_DE)) {
            writeCell.addCell(row, ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle, lineOfLedger.toString(), null,
                    "le grand livre");
            writeCell.addCell(row, ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte de la contrepartie", "le grand livre");
            writeCell.addCell(row, ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                    lineOfLedger.toString(), "le libellé de la contrepartie", "le grand livre");
        } else {
            writeCell.addCellEmpty(ID_JOURNAL_OF_LEDGER, ID_CHECK_OF_LEDGER, row, cellStyle);
        }
        writeCell.addCell(row, ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle, lineOfLedger.toString(), null,
                "le grand livre");
        writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle, lineOfLedger.toString(),
                "le libellé", "le grand livre");

        Cell debitCell = writeCell.addDebitCell(lineOfLedger, row, cellStyleAmount);
        Cell creditCell = writeCell.addCreditCell(lineOfLedger, row, cellStyleAmount);
        writeCell.addSoldeCell(lineOfLedger, row, cellStyleAmount, creditCell, debitCell);

        if (verif) {
            writeOutil.addVerifCells(lineOfLedger, row, cellStyle, pathDirectoryInvoice);
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
            writeCell.addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfLedger.account().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte", place);
            writeCell.addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfLedger.account().label(), cellStyle, lineOfLedger.toString(),
                    "le libellé de compte", place);
            writeCell.addCell(row, ID_DOCUMENT_OF_LEDGER, lineOfLedger.document(), cellStyle, lineOfLedger.toString(),
                    "la piéce", place);
            writeCell.addCell(row, ID_DATE_OF_LEDGER, lineOfLedger.date(), cellStyle, lineOfLedger.toString(),
                    "la date", place);
            writeCell.addCell(row, ID_JOURNAL_OF_LEDGER, lineOfLedger.journal(), cellStyle, lineOfLedger.toString(), null,
                    place);
            writeCell.addCell(row, ID_COUNTERPART_NUMBER_OF_LEDGER, lineOfLedger.accountCounterpart().account(), cellStyle,
                    lineOfLedger.toString(), "le numéro de compte de la contrepartie", place);
            writeCell.addCell(row, ID_COUNTERPART_LABEL_OF_LEDGER, lineOfLedger.accountCounterpart().label(), cellStyle,
                    lineOfLedger.toString(), "le libellé de la contrepartie", place);
            writeCell.addCell(row, ID_CHECK_OF_LEDGER, lineOfLedger.checkNumber(), cellStyle, lineOfLedger.toString(),
                    null, place);
            writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfLedger.label(), cellStyle, lineOfLedger.toString(),
                    "le libellé", place);
            writeCell.addDebitCell(lineOfLedger, row, cellStyleAmount);
            writeCell.addCreditCell(lineOfLedger, row, cellStyleAmount);
        }
        if (bankLine != null) {
            String place = "le relevé de compte banque";
            writeCell.addCell(row, ID_MONTH_OF_SATEMENT_OF_RECONCILIATION, bankLine.year() + "-" + bankLine.mounth(),
                    cellStyle, bankLine.toString(), "le mois et l'année", place);
            writeCell.addCell(row, ID_ACOUNT_NUMBER_OF_RECONCILIATION, bankLine.account().account(), cellStyle,
                    bankLine.toString(), "le numéro du compte", place);
            writeCell.addCell(row, ID_ACOUNT_LABEL_OF_RECONCILIATION, bankLine.account().label(), cellStyle, bankLine.toString(),
                    "le libellé du compte", place);
            writeCell.addCell(row, ID_OPERATION_DATE_OF_RECONCILIATION, bankLine.operationDate().format(DATE_FORMATTER),
                    cellStyle, bankLine.toString(), "la date de l'opération", place);
            writeCell.addCell(row, ID_VALUE_DATE_OF_RECONCILIATION, bankLine.valueDate().format(DATE_FORMATTER), cellStyle,
                    bankLine.toString(), "la date de valeur", place);
            writeCell.addCell(row, ID_LABEL_OF_RECONCILIATION, bankLine.label(), cellStyle, bankLine.toString(),
                    "le libellé", place);

            Cell debitReleveCell = row.createCell(ID_DEBIT_OF_RECONCILIATION);
            debitReleveCell.setCellValue(bankLine.debit());
            debitReleveCell.setCellStyle(cellStyleAmount);

            Cell creditReleveCell = row.createCell(ID_CREDIT_OF_RECONCILIATION);
            creditReleveCell.setCellValue(bankLine.credit());
            creditReleveCell.setCellStyle(cellStyleAmount);
        }
        writeCell.addCell(row, ID_COMMENT_OF_RECONCILIATION, message, cellStyle, null, null, null);
    }

    public void getTotalBuilding(TotalBuilding lineOfTotalBuildingInLedger, Row row, List<Integer> lineTotals) {

        CellStyle styleTotal = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle styleTotalAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        writeCell.addCellEmpty(ID_ACOUNT_NUMBER_OF_LEDGER, ID_LABEL_OF_LEDGER, row, styleTotal);
        writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfTotalBuildingInLedger.label(), styleTotal,
                lineOfTotalBuildingInLedger.toString(), "le libellé", "le grand livre");

        Cell debitCell = writeCell.addCellAmountOfTotalBuildingInLedger(row, ID_DEBIT_OF_LEDGER, lineTotals, styleTotalAmount);
        Cell creditCell = writeCell.addCellAmountOfTotalBuildingInLedger(row, ID_CREDIT_OF_LEDGER, lineTotals, styleTotalAmount);

        writeCell.addCellBalanceOfTotalInLedger(row, debitCell, creditCell, styleTotalAmount);
    }

    public void getTotalAccount(TotalAccount lineOfTotalAccountInLedger, Row row, int lastRowNumTotal) {

        CellStyle cellStyle = writeCellStyle.getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle cellStyleAmount = writeCellStyle.getCellStyleTotalAmount(row.getSheet().getWorkbook());

        writeCell.addCell(row, ID_ACOUNT_NUMBER_OF_LEDGER, lineOfTotalAccountInLedger.account().account(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le numéro de compte", "le grand livre");
        writeCell.addCell(row, ID_ACOUNT_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.account().label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé du compte", "le grand livre");
        writeCell.addCellEmpty(ID_DOCUMENT_OF_LEDGER, ID_LABEL_OF_LEDGER, row, cellStyle);
        writeCell.addCell(row, ID_LABEL_OF_LEDGER, lineOfTotalAccountInLedger.label(), cellStyle,
                lineOfTotalAccountInLedger.toString(), "le libellé", "le grand livre");

        Cell debitCell = writeCell.addCellDebitOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell creditCell = writeCell.addCellCreditOfTotalAccountInLedger(row, lastRowNumTotal, cellStyleAmount);
        Cell soldeCell = writeCell.addCellBalanceOfTotalInLedger(row, debitCell, creditCell, cellStyleAmount);

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

    public void getCellsEnteteGrandLivre(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_GRAND_LIVRE) {
            writeCell.addCell(headerRow, index++, label, styleHeader, "", null, null);
        }
    }

    public void getCellsEnteteListeDesDepenses(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle cellStyleHeader = writeCellStyle.getCellStyleEntete(sheet.getWorkbook().createCellStyle());
        for (String label : NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES) {
            writeCell.addCell(headerRow, index++, label, cellStyleHeader, "", null, null);
        }
    }

    public void getCellsEnteteEtatRapprochement(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT) {
            writeCell.addCell(headerRow, index++, label, writeCellStyle.getCellStyleEntete(styleHeader), "", null, null);
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
        if (writeOutil.isDouble(line.document())) {
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

        String message = writeOutil.getMessageFindDocument(line.document(), link, pathDirectoryInvoice);
        if (message.startsWith(IMPOSSIBLE_DE_TROUVER_LA_PIECE)) {
            message = message + " avec ce libellé " + line.label();
        }
        cell.setCellValue(message);
        cell.setCellStyle(styleColor);
    }


}
