package life.light.write;

import life.light.type.*;
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

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    public static final String AMOUNT_FORMATTER = "# ### ##0.00 €;[red]# ### ##0.00 €";
    private static final Logger LOGGER = LogManager.getLogger();
    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_WHITE = new XSSFColor(new java.awt.Color(255, 255, 255), null);
    private static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    private static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);
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
            operationDateReleveCell.setCellValue(bankLine.operationDate().format(DATE_FORMATTER));
            addlineBlue(getCellStyleAlignmentLeft(style), operationDateReleveCell);
            cellNum++;
            Cell valueDateReleveCell = row.createCell(cellNum);
            valueDateReleveCell.setCellValue(bankLine.valueDate().format(DATE_FORMATTER));
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
                message = "Impossible de trouver la pièce " + document
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

    public void getCellsEnteteGrandLivre(Sheet sheet, CellStyle styleHeader) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        for (String label : NOM_ENTETE_COLONNE_GRAND_LIVRE) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(getCellStyleEntete(styleHeader));
        }
    }

    public void getCellsEnteteListeDesDepenses(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle cellStyleHeader = getCellStyleEntete(sheet.getWorkbook().createCellStyle());
        for (String label : NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(cellStyleHeader);
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

    public CellStyle getCellStyleTotalAmount(Workbook workbook) {
        CellStyle style = getCellStyleAmount(workbook);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    public CellStyle getCellStyleAmount(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(AMOUNT_FORMATTER));
        return style;
    }

    public CellStyle getCellStyleTotal(Workbook workbook) {
        CellStyle styleTotal = workbook.createCellStyle();
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

    public void getLineOfExpenseKey(LineOfExpenseKey line, Row row, CellStyle styleTotal) {
        Cell cell;
        for (int index = 0; index < NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES.length; index++) {
            cell = row.createCell(index);
            if (index == ID_LABEL_OF_LIST_OF_EXPENSES) {
                cell.setCellValue(line.label() + " : " + line.key() + " " + line.value());
            }
            cell.setCellStyle(styleTotal);
        }
    }

    public void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row, CellStyle styleTotal, CellStyle styleTotalAmount) {
        Cell cell;
        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(styleTotal);
        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(styleTotal);
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
            cell.setCellStyle(styleTotal);
        }
        if (!line.amount().isEmpty()) {
            cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.amount()));
            cell.setCellStyle(styleTotalAmount);
        }
        if (!line.deduction().isEmpty()) {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.deduction()));
            cell.setCellStyle(styleTotalAmount);
        }
        if (!line.recovery().isEmpty()) {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.recovery()));
            cell.setCellStyle(styleTotalAmount);
        }
    }

    public void getLineOfExpense(LineOfExpense line, Row row, CellStyle styleColor, CellStyle styleAmountColor, String pathDirectoryInvoice) {
        CreationHelper createHelper = row.getSheet().getWorkbook().getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);

        Cell cell;

        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        cell.setCellValue(line.document());
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
        if (message.startsWith("Impossible de trouver la pièce")) {
            message = message + " avec ce libellé " + line.label();
        }
        cell.setCellValue(message);
        cell.setCellStyle(styleColor);
    }

    public CellStyle getCellStyleWhite(Workbook workbook) {
        CellStyle styleWhite = workbook.createCellStyle();
        styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        styleWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleWhite;
    }

    public CellStyle getCellStyleBlue(Workbook workbook) {
        CellStyle styleBlue = workbook.createCellStyle();
        styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        styleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleBlue;
    }

    public CellStyle getCellStyleAmountWhite(Workbook workbook) {
        CellStyle styleAmountWhite = getCellStyleAmount(workbook);
        styleAmountWhite.setFillBackgroundColor(BACKGROUND_COLOR_WHITE);
        return styleAmountWhite;
    }

    CellStyle getCellStyleAmountBlue(Workbook workbook) {
        CellStyle styleAmountBlue = getCellStyleAmount(workbook);
        styleAmountBlue.setFillBackgroundColor(BACKGROUND_COLOR_BLUE);
        return styleAmountBlue;
    }
}