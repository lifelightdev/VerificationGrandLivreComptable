package life.light.write;

import life.light.type.BankLine;
import life.light.type.Line;
import life.light.type.TotalAccount;
import life.light.type.TotalBuilding;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class OutilWrite {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_WHITE = new XSSFColor(new java.awt.Color(255, 255, 255), null);
    private static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    private static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);

    protected static final String[] NOM_ENTETE_COLONNE_GRAND_LIVRE = {"Compte", "Intitulé du compte", "Pièce", "Date",
            "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "Solde (Calculé)", "Vérification", "Commentaire"};
    protected static final String[] NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT = {"Compte", "Intitulé du compte", "Pièce",
            "Date", "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "----",
            "Mois du relevé", "Compte", "Intitulé du compte", "Date de l'opération", "Date de valeur", "Libellé",
            "Débit", "Crédit", "Commentaire"};
    private static final String REPORT_DE = "Report de";

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

    public void getTotalBuilding(TotalBuilding grandLivre, Row row, CellStyle styleTotal, CellStyle styleTotalAmount,
                                 Workbook workbook, List<Integer> lineTotals) {
        Cell cell;
        for (int idCell = 0; idCell < 9; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(styleTotal);
        }
        int cellNum = 8;
        if (!grandLivre.label().isEmpty()) {
            cell = row.createCell(cellNum++);
            cell.setCellValue(grandLivre.label());
            cell.setCellStyle(styleTotal);
        }
        Cell debitCell = row.createCell(cellNum++);
        StringBuilder sumDebit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumDebit.append(CellReference.convertNumToColString(debitCell.getColumnIndex())).append(numRow).append("+");
        }
        debitCell.setCellFormula(sumDebit.substring(0, sumDebit.lastIndexOf("+")));
        workbook.setForceFormulaRecalculation(true);
        debitCell.setCellStyle(styleTotalAmount);

        Cell creditCell = row.createCell(cellNum++);
        StringBuilder sumCredit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumCredit.append(CellReference.convertNumToColString(creditCell.getColumnIndex())).append(numRow).append("+");
        }
        creditCell.setCellFormula(sumCredit.substring(0, sumCredit.lastIndexOf("+")));
        workbook.setForceFormulaRecalculation(true);
        creditCell.setCellStyle(styleTotalAmount);

        Cell soldeCell = row.createCell(cellNum);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(styleTotalAmount);
        workbook.setForceFormulaRecalculation(true);
    }

    public void getTotalAccount(TotalAccount grandLivre, Row row, CellStyle styleTotal, CellStyle styleTotalAmount, Workbook workbook, int lastRowNumTotal) {
        Cell cell;
        if (!grandLivre.account().account().isEmpty()) {
            cell = row.createCell(0);
            cell.setCellValue(grandLivre.account().account());
            if (isDouble(grandLivre.account().account())) {
                double account = Double.parseDouble(grandLivre.account().account());
                cell.setCellValue(account);
            }
            cell.setCellStyle(styleTotal);
        }
        if (!grandLivre.account().label().isEmpty()) {
            cell = row.createCell(1);
            cell.setCellValue(grandLivre.account().label());
            cell.setCellStyle(styleTotal);
        }
        for (int idCell = 2; idCell < 8; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(styleTotal);
        }
        int cellNum = 8;
        if (!grandLivre.label().isEmpty()) {
            cell = row.createCell(cellNum++);
            cell.setCellValue(grandLivre.label());
            cell.setCellStyle(styleTotal);
        }
        Cell debitCell = row.createCell(cellNum++);
        CellAddress debitCellAddressFirst = new CellAddress(lastRowNumTotal + 1, debitCell.getAddress().getColumn());
        CellAddress debitCellAddressEnd = new CellAddress(debitCell.getAddress().getRow() - 1, debitCell.getAddress().getColumn());
        debitCell.setCellFormula("SUM(" + debitCellAddressFirst + ":" + debitCellAddressEnd + ")");
        workbook.setForceFormulaRecalculation(true);
        debitCell.setCellStyle(styleTotalAmount);

        Cell creditCell = row.createCell(cellNum++);
        creditCell.setCellStyle(styleTotalAmount);
        CellAddress creditCellAddressFirst = new CellAddress(lastRowNumTotal + 1, creditCell.getAddress().getColumn());
        CellAddress creditCellAddressEnd = new CellAddress(creditCell.getAddress().getRow() - 1, creditCell.getAddress().getColumn());
        creditCell.setCellFormula("SUM(" + creditCellAddressFirst + ":" + creditCellAddressEnd + ")");

        Cell soldeCell = row.createCell(cellNum++);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(styleTotalAmount);
        workbook.setForceFormulaRecalculation(true);

        Cell cellVerif = row.createCell(cellNum++);
        String amount = grandLivre.label()
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
        workbook.setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        double solde = (double) Math.round(evaluator.evaluate(soldeCell).getNumberValue() * 100) / 100;
        double debitExcel = (double) Math.round(evaluator.evaluate(debitCell).getNumberValue() * 100) / 100;
        double debitGrandLivre = Double.parseDouble(grandLivre.debit());
        double creditExcel = (double) Math.round(evaluator.evaluate(creditCell).getNumberValue() * 100) / 100;
        double creditGrandLivre = Double.parseDouble(grandLivre.credit());
        if (Double.parseDouble(amount) == solde) {
            if (debitExcel == debitGrandLivre) {
                if (creditExcel == creditGrandLivre) {
                    verif = "OK";
                    cellVerif.setCellStyle(styleTotalAmount);
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
            cellVerif.setCellStyle(getCellStyleVerifRed(styleTotal));
            LOGGER.info("Vérification du total du compte [{}] Solde PDF [{}] Solde Excel [{}] débit PDF [{}] débit Excel [{}] crédit PDF [{}] crédit Excel [{}] ligne [{}]",
                    grandLivre.account().account(), amount, solde,
                    debitGrandLivre, debitExcel,
                    creditGrandLivre, creditExcel,
                    grandLivre);
        }
        cellVerif.setCellValue(verif);

        Cell cellMessqage = row.createCell(cellNum);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount) + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        //cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(styleTotalAmount);
    }

    public void getLineGrandLivre(Line grandLivre, Row row, CellStyle style, CellStyle styleAmount, Workbook workbook, boolean verif, String pathDirectoryInvoice) {
        int cellNum = 0;

        cellNum = addAccountCell(grandLivre, row, style, cellNum);
        cellNum = addDocumentCell(grandLivre, row, style, cellNum);
        cellNum = addDateCell(grandLivre, row, style, cellNum);
        cellNum = addJournalCell(grandLivre, row, style, cellNum);
        cellNum = addCounterPartCell(grandLivre, row, style, cellNum);
        cellNum = addCheckNumberCell(grandLivre, row, style, cellNum);
        cellNum = addLabelCell(grandLivre, row, style, cellNum);

        Cell debitCell = addDebitCell(grandLivre, row, styleAmount, cellNum);
        cellNum++;

        Cell creditCell = addCreditCell(grandLivre, row, styleAmount, cellNum);
        cellNum++;

        cellNum = addSoldeCell(grandLivre, row, styleAmount, cellNum, creditCell, debitCell);

        if (verif) {
            addVerifCells(grandLivre, row, style, workbook, cellNum, pathDirectoryInvoice);
        }
    }

    public void getLineEtatRapprochement(Line grandLivre, Row row, CellStyle style, CellStyle styleAmount, BankLine bankLine, String message) {
        int cellNum = 0;
        if (grandLivre != null) {
            //LOGGER.info("Dans l'onglet " + row.getSheet().getSheetName() + " ajout de la ligne du grand livre " + grandLivre.account().account() + " label " + grandLivre.label() + " debit " + grandLivre.debit() + " credit " + grandLivre.credit());
            cellNum = addAccountCell(grandLivre, row, style, cellNum);
            cellNum = addDocumentCell(grandLivre, row, style, cellNum);
            cellNum = addDateCell(grandLivre, row, style, cellNum);
            cellNum = addJournalCell(grandLivre, row, style, cellNum);
            cellNum = addCounterPartCell(grandLivre, row, style, cellNum);
            cellNum = addCheckNumberCell(grandLivre, row, style, cellNum);
            cellNum = addLabelCell(grandLivre, row, style, cellNum);
            addDebitCell(grandLivre, row, styleAmount, cellNum);
            cellNum++;
            addCreditCell(grandLivre, row, styleAmount, cellNum);
            cellNum++;
        } else {
            cellNum = 11;
        }
        cellNum++;
        if (bankLine != null) {
            //LOGGER.info(("Dans l'onglet " + row.getSheet().getSheetName() + " ajout de la ligne du relever de compte " + bankLine.operationDate() + " label " + bankLine.label() + " debit " + bankLine.debit() + " credit " + bankLine.credit()));
            Cell dateReleveCell = row.createCell(cellNum);
            dateReleveCell.setCellValue(bankLine.year() + "-" + bankLine.mounth());
            addlineBlue(getCellStyleAlignmentLeft(style), dateReleveCell);
            cellNum++;
            Cell accountReleveCell = row.createCell(cellNum);
            accountReleveCell.setCellValue(Double.parseDouble(bankLine.account().account()));
            addlineBlue(getCellStyleAlignmentLeft(style), accountReleveCell);
            cellNum++;
            Cell labelAccountReleveCell = row.createCell(cellNum);
            labelAccountReleveCell.setCellValue(bankLine.account().label());
            addlineBlue(getCellStyleAlignmentLeft(style), labelAccountReleveCell);
            cellNum++;
            Cell operationDateReleveCell = row.createCell(cellNum);
            DateTimeFormatter formateur = DateTimeFormatter.ofPattern("dd/MM/yyyy");
            operationDateReleveCell.setCellValue(bankLine.operationDate().format(formateur));
            addlineBlue(getCellStyleAlignmentLeft(style), operationDateReleveCell);
            cellNum++;
            Cell valueDateReleveCell = row.createCell(cellNum);
            valueDateReleveCell.setCellValue(bankLine.valueDate().format(formateur));
            addlineBlue(getCellStyleAlignmentLeft(style), valueDateReleveCell);
            cellNum++;
            Cell labelReleveCell = row.createCell(cellNum);
            labelReleveCell.setCellValue(bankLine.label());
            addlineBlue(getCellStyleAlignmentLeft(style), labelReleveCell);
            cellNum++;
            Cell debitReleveCell = row.createCell(cellNum);
            debitReleveCell.setCellValue(bankLine.debit());
            addlineBlue(styleAmount, debitReleveCell);
            cellNum++;
            Cell creditReleveCell = row.createCell(cellNum);
            creditReleveCell.setCellValue(bankLine.credit());
            addlineBlue(styleAmount, creditReleveCell);
        }
        Cell messageCell = row.createCell(20);
        messageCell.setCellValue(message);
        addlineBlue(styleAmount, messageCell);
    }

    private int addVerifCells(Line grandLivre, Row row, CellStyle style, Workbook workbook, int cellNum, String pathDirectoryInvoice) {
        String message = "";
        CreationHelper createHelper = workbook.getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
        Cell verifCell = row.createCell(cellNum);
        if (grandLivre.label().startsWith(REPORT_DE)) {
            message = getMessageVerifLineReport(grandLivre, verifCell, style, message);
        } else {
            if (grandLivre.account().account().startsWith(CLASSE_6)
                    || (grandLivre.account().account().startsWith("401")) && (grandLivre.accountCounterpart().account().startsWith(CLASSE_6))) {
                message = getMessageFindDocument(grandLivre, verifCell, message, link, pathDirectoryInvoice);
            }
            if (grandLivre.accountCounterpart() == null) {
                verifCell.setCellValue(KO);
                message = "Il n'y a pas de Contrepartie";
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre);
            }
            if (grandLivre.debit().isEmpty() && grandLivre.credit().isEmpty()) {
                verifCell.setCellValue(KO);
                message = "Il n'y a aucun montant";
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre);
            }
        }
        if (verifCell.getStringCellValue().equals(KO)) {
            verifCell.setCellStyle(getCellStyleVerifRed(style));
        } else {
            addlineBlue(style, verifCell);
        }

        cellNum++;
        Cell messageCell = row.createCell(cellNum);
        if (link.getAddress() != null) {
            messageCell.setHyperlink(link);
        }
        messageCell.setCellValue(message.trim());
        addlineBlue(style, messageCell);
        return cellNum;
    }

    private String getMessageFindDocument(Line grandLivre, Cell verifCell, String message, Hyperlink link, String thePathDirectoryInvoice) {
        File pathDirectoryInvoice = new File(thePathDirectoryInvoice);
        File[] files = pathDirectoryInvoice.listFiles();
        if (null != files) {
            boolean find = false;
            for (File fichier : files) {
                if (fichier.isFile()) {
                    if (fichier.getName().contains(grandLivre.document())) {
                        verifCell.setCellValue("OK");
                        find = true;
                        message = fichier.getAbsoluteFile().toString().replace("F:", "D:");
                        link.setAddress(fichier.toURI().toString().replace("F:", "D:"));
                        break;
                    }
                } else if (fichier.isDirectory()) {
                    File[] sousDossier = fichier.listFiles();
                    if (null != sousDossier) {
                        for (File fichierDuSousDossier : sousDossier) {
                            if (fichierDuSousDossier.isFile()) {
                                if (fichierDuSousDossier.getName().contains(grandLivre.document())) {
                                    verifCell.setCellValue("OK");
                                    find = true;
                                    message = fichierDuSousDossier.getAbsoluteFile().toString().replace("F:", "D:");
                                    link.setAddress(fichierDuSousDossier.toURI().toString().replace("F:", "D:"));
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            if (!find) {
                message = "Impossible de trouver la pièce";
                LOGGER.info("{} : {} dans le dossier : {} sur le compte {} libelle de l'opération {}",
                        message, grandLivre.document(), pathDirectoryInvoice, grandLivre.account().account(), grandLivre.label());
            }
        }
        return message;
    }

    private String getMessageVerifLineReport(Line grandLivre, Cell verifCell, CellStyle style, String message) {
        getCellStyleVerifRed(style);
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
                addlineBlue(style, verifCell);
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

    private int addSoldeCell(Line grandLivre, Row row, CellStyle style, int cellNum, Cell creditCell, Cell debitCell) {
        Cell soldeCell = row.createCell(cellNum);
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
        addlineBlue(style, soldeCell);
        cellNum++;
        return cellNum;
    }

    private Cell addCreditCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell creditCell = row.createCell(cellNum);
        creditCell.setCellValue(grandLivre.credit());
        if (isDouble(grandLivre.credit())) {
            creditCell.setCellValue(Double.parseDouble(grandLivre.credit()));
        } else {
            creditCell.setCellValue(0);
        }
        addlineBlue(style, creditCell);
        return creditCell;
    }

    private Cell addDebitCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell debitCell = row.createCell(cellNum);
        debitCell.setCellValue(grandLivre.debit());
        if (isDouble(grandLivre.debit())) {
            debitCell.setCellValue(Double.parseDouble(grandLivre.debit()));
        } else {
            debitCell.setCellValue(0D);
        }
        addlineBlue(style, debitCell);
        return debitCell;
    }

    private int addLabelCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell labelCell = row.createCell(cellNum);
        if (!grandLivre.label().isEmpty()) {
            labelCell.setCellValue(grandLivre.label());
        } else {
            LOGGER.error("Le libellé est absent sur la ligne : {}", grandLivre);
        }
        addlineBlue(getCellStyleAlignmentLeft(style), labelCell);
        cellNum++;
        return cellNum;
    }

    private int addCheckNumberCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell checkNumberCell = row.createCell(cellNum);
        if (isDouble(grandLivre.checkNumber())) {
            checkNumberCell.setCellValue(Double.parseDouble(grandLivre.checkNumber()));
        } else {
            checkNumberCell.setCellValue(grandLivre.checkNumber());
        }
        addlineBlue(getCellStyleAlignmentLeft(style), checkNumberCell);
        cellNum++;
        return cellNum;
    }

    private int addCounterPartCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell counterPartCell = row.createCell(cellNum++);
        if (grandLivre.accountCounterpart() != null) {
            if (isDouble(grandLivre.accountCounterpart().account())) {
                counterPartCell.setCellValue(Double.parseDouble(grandLivre.accountCounterpart().account()));
            } else {
                counterPartCell.setCellValue(grandLivre.accountCounterpart().account());
            }
            Cell labelCounterPartCell = row.createCell(cellNum);
            labelCounterPartCell.setCellValue(grandLivre.accountCounterpart().label());
            addlineBlue(getCellStyleAlignmentLeft(style), labelCounterPartCell);

        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("La contre partie est absente sur la ligne : {}", grandLivre);
        }
        addlineBlue(getCellStyleAlignmentLeft(style), counterPartCell);
        cellNum++;
        return cellNum;
    }

    private int addJournalCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell journalCell = row.createCell(cellNum);
        journalCell.setCellValue(grandLivre.journal());
        if (!grandLivre.journal().isEmpty()) {
            if (isDouble(grandLivre.journal())) {
                journalCell.setCellValue(Double.parseDouble(grandLivre.journal()));
            } else {
                journalCell.setCellValue(grandLivre.journal());
            }
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("Le journal est absent sur la ligne : {}", grandLivre);
        }
        addlineBlue(getCellStyleAlignmentLeft(style), journalCell);
        cellNum++;
        return cellNum;
    }

    private int addDateCell(Line grandLivre, Row row, CellStyle style, int cellNum) {
        Cell dateCell = row.createCell(cellNum);
        dateCell.setCellValue(grandLivre.date());
        addlineBlue(getCellStyleAlignmentLeft(style), dateCell);
        cellNum++;
        return cellNum;
    }

    private int addDocumentCell(Line grandLivre, Row row, CellStyle cellStyle, int cellNum) {
        Cell documentCell = row.createCell(cellNum);
        documentCell.setCellValue(grandLivre.document());
        if (isDouble(grandLivre.document())) {
            double document = Double.parseDouble(grandLivre.document());
            documentCell.setCellValue(document);
        } else if ((!grandLivre.label().contains(REPORT_DE)) && grandLivre.document().isEmpty()) {
            LOGGER.error("Le numéro de piéce est absente sur la ligne : {}", grandLivre);
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("Le numéro de piéce n'est pas numérique ({}) sur la ligne : {}", grandLivre.document(), grandLivre);
        }
        addlineBlue(getCellStyleAlignmentLeft(cellStyle), documentCell);
        cellNum++;
        return cellNum;
    }

    private int addAccountCell(Line grandLivre, Row row, CellStyle cellStyle, int cellNum) {
        String numAccount = grandLivre.account().account();
        boolean isNotEmptyAccount = !numAccount.isEmpty();
        if (isNotEmptyAccount) {
            Cell accountNumberCell = getAccountNumberCell(numAccount, row, cellNum++);
            addlineBlue(getCellStyleAlignmentLeft(cellStyle), accountNumberCell);

            Cell accountLabelCell = row.createCell(cellNum++);
            accountLabelCell.setCellValue(grandLivre.account().label());
            addlineBlue(getCellStyleAlignmentLeft(cellStyle), accountLabelCell);
        } else {
            LOGGER.error("Il manque le numéro de compte sur cette ligne : {}", grandLivre);
        }
        return cellNum;
    }

    Cell getAccountNumberCell(String accountNumber, Row row, int numColAccount) {
        Cell accountNumberCell = row.createCell(numColAccount);
        // Pour ne pas avoir d'erreur de formatage dans excel
        if (isDouble(accountNumber)) {
            accountNumberCell.setCellValue(Double.parseDouble(accountNumber));
        } else {
            accountNumberCell.setCellValue(accountNumber);
        }
        return accountNumberCell;
    }

    private CellStyle getCellStyleVerifRed(CellStyle style) {
        style.setFillForegroundColor(BACKGROUND_COLOR_RED);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private void addlineBlue(CellStyle cellStyle, Cell cell) {
        if (cell.getRowIndex() % 2 == 0) {
            cellStyle.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        } else {
            cellStyle.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        }
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
    }

    public void getCellsEnteteGrandLivre(Sheet sheet, CellStyle styleHeader) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        for (String label : NOM_ENTETE_COLONNE_GRAND_LIVRE) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(getCellStyleEntete(styleHeader));
        }
    }

    public void getCellsEnteteEtatRapprochement(Sheet sheet, CellStyle styleHeader) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        for (String label : NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(getCellStyleEntete(styleHeader));
        }
    }

    public CellStyle getCellStyleTotalAmount(CellStyle style, Short dataAmount) {
        style = getCellStyleAmount(style, dataAmount);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    public CellStyle getCellStyleAmount(CellStyle style, Short dataAmount) {
        style.setDataFormat(dataAmount);
        return style;
    }

    public CellStyle getCellStyleTotal(CellStyle styleTotal) {
        styleTotal.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleTotal;
    }

    private CellStyle getCellStyleAlignmentLeft(CellStyle style) {
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    private CellStyle getCellStyleEntete(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private boolean isDouble(String str) {
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
}