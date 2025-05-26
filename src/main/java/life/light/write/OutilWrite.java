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
import java.util.Map;
import java.util.TreeMap;

public class OutilWrite {

    private static final Logger LOGGER = LogManager.getLogger();

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");
    public static final String AMOUNT_FORMATTER = "# ### ##0.00 €;[red]# ### ##0.00 €";

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

    public void getTotalBuilding(TotalBuilding grandLivre, Row row, List<Integer> lineTotals) {
        Cell cell;
        CellStyle styleTotal = getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle styleTotalAmount = getCellStyleTotalAmount(row.getSheet().getWorkbook());
        for (int idCell = ID_ACOUNT_NUMBER_OF_LEDGER; idCell < ID_LABEL_OF_LEDGER; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(styleTotal);
        }

        cell = row.createCell(ID_LABEL_OF_LEDGER);
        cell.setCellValue(grandLivre.label());
        cell.setCellStyle(styleTotal);

        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        StringBuilder sumDebit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumDebit.append(CellReference.convertNumToColString(debitCell.getColumnIndex())).append(numRow).append("+");
        }
        debitCell.setCellFormula(sumDebit.substring(0, sumDebit.lastIndexOf("+")));
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(styleTotalAmount);

        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        StringBuilder sumCredit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumCredit.append(CellReference.convertNumToColString(creditCell.getColumnIndex())).append(numRow).append("+");
        }
        creditCell.setCellFormula(sumCredit.substring(0, sumCredit.lastIndexOf("+")));
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        creditCell.setCellStyle(styleTotalAmount);

        Cell soldeCell = row.createCell(ID_BALANCE_OF_LEDGER);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(styleTotalAmount);
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
    }

    public void getTotalAccount(TotalAccount grandLivre, Row row, int lastRowNumTotal) {
        Cell cell;
        CellStyle cellStyle = getCellStyleTotal(row.getSheet().getWorkbook());
        CellStyle cellStyleAmount = getCellStyleTotalAmount(row.getSheet().getWorkbook());
        if (!grandLivre.account().account().isEmpty()) {
            cell = row.createCell(ID_ACOUNT_NUMBER_OF_LEDGER);
            cell.setCellValue(grandLivre.account().account());
            if (isDouble(grandLivre.account().account())) {
                double account = Double.parseDouble(grandLivre.account().account());
                cell.setCellValue(account);
            }
            cell.setCellStyle(cellStyle);
        }
        if (!grandLivre.account().label().isEmpty()) {
            cell = row.createCell(ID_ACOUNT_LABEL_OF_LEDGER);
            cell.setCellValue(grandLivre.account().label());
            cell.setCellStyle(cellStyle);
        }
        for (int idCell = ID_DOCUMENT_OF_LEDGER; idCell < ID_LABEL_OF_LEDGER; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(cellStyle);
        }

        cell = row.createCell(ID_LABEL_OF_LEDGER);
        cell.setCellValue(grandLivre.label());
        cell.setCellStyle(cellStyle);

        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        CellAddress debitCellAddressFirst = new CellAddress(lastRowNumTotal + 1, debitCell.getAddress().getColumn());
        CellAddress debitCellAddressEnd = new CellAddress(debitCell.getAddress().getRow() - 1, debitCell.getAddress().getColumn());
        debitCell.setCellFormula("SUM(" + debitCellAddressFirst + ":" + debitCellAddressEnd + ")");
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(cellStyleAmount);

        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        creditCell.setCellStyle(cellStyleAmount);
        CellAddress creditCellAddressFirst = new CellAddress(lastRowNumTotal + 1, creditCell.getAddress().getColumn());
        CellAddress creditCellAddressEnd = new CellAddress(creditCell.getAddress().getRow() - 1, creditCell.getAddress().getColumn());
        creditCell.setCellFormula("SUM(" + creditCellAddressFirst + ":" + creditCellAddressEnd + ")");

        Cell soldeCell = row.createCell(ID_BALANCE_OF_LEDGER);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(cellStyleAmount);
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);

        Cell cellVerif = row.createCell(ID_VERIFFICATION_OF_LEDGER);
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
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        double solde = (double) Math.round(evaluator.evaluate(soldeCell).getNumberValue() * 100) / 100;
        double debitExcel = (double) Math.round(evaluator.evaluate(debitCell).getNumberValue() * 100) / 100;
        double debitGrandLivre = Double.parseDouble(grandLivre.debit());
        double creditExcel = (double) Math.round(evaluator.evaluate(creditCell).getNumberValue() * 100) / 100;
        double creditGrandLivre = Double.parseDouble(grandLivre.credit());
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
            cellVerif.setCellStyle(getCellStyleVerifRed(cellStyle));
            LOGGER.info("Vérification du total du compte [{}] Solde PDF [{}] Solde Excel [{}] débit PDF [{}] débit Excel [{}] crédit PDF [{}] crédit Excel [{}] ligne [{}]",
                    grandLivre.account().account(), amount, solde,
                    debitGrandLivre, debitExcel,
                    creditGrandLivre, creditExcel,
                    grandLivre);
        }
        cellVerif.setCellValue(verif);

        Cell cellMessqage = row.createCell(ID_COMMENT_OF_LEDGER);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount) + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        //cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(cellStyle);
    }

    public void getLineGrandLivre(Line grandLivre, Row row, boolean verif, String pathDirectoryInvoice) {
        CellStyle cellStyle;
        CellStyle cellStyleAmount;
        if (row.getRowNum() % 2 == 0) {
            cellStyle = getCellStyleWhite(row.getSheet().getWorkbook());
            cellStyleAmount = getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            cellStyle = getCellStyleBlue(row.getSheet().getWorkbook());
            cellStyleAmount = getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }

        addAccountCell(grandLivre, row, cellStyle);
        addDocumentCell(grandLivre, row, cellStyle);
        addDateCell(grandLivre, row, cellStyle);
        addJournalCell(grandLivre, row, cellStyle);
        addCounterPartCell(grandLivre, row, cellStyle);
        addCheckNumberCell(grandLivre, row, cellStyle);
        addLabelCell(grandLivre, row, cellStyle);

        Cell debitCell = addDebitCell(grandLivre, row, cellStyleAmount);
        Cell creditCell = addCreditCell(grandLivre, row, cellStyleAmount);
        addSoldeCell(grandLivre, row, cellStyleAmount, creditCell, debitCell);

        if (verif) {
            addVerifCells(grandLivre, row, cellStyle, pathDirectoryInvoice);
        }
    }

    public void getLineEtatRapprochement(Line grandLivre, Row row, BankLine bankLine, String message) {
        CellStyle cellStyle;
        CellStyle cellStyleAmount;
        if (row.getRowNum() % 2 == 0) {
            cellStyle = getCellStyleWhite(row.getSheet().getWorkbook());
            cellStyleAmount = getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            cellStyle = getCellStyleBlue(row.getSheet().getWorkbook());
            cellStyleAmount = getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }
        if (grandLivre != null) {
            addAccountCell(grandLivre, row, cellStyle);
            addDocumentCell(grandLivre, row, cellStyle);
            addDateCell(grandLivre, row, cellStyle);
            addJournalCell(grandLivre, row, cellStyle);
            addCounterPartCell(grandLivre, row, cellStyle);
            addCheckNumberCell(grandLivre, row, cellStyle);
            addLabelCell(grandLivre, row, cellStyle);
            addDebitCell(grandLivre, row, cellStyleAmount);
            addCreditCell(grandLivre, row, cellStyleAmount);
        }
        if (bankLine != null) {
            Cell dateReleveCell = row.createCell(ID_MONTH_OF_SATEMENT_OF_RECONCILIATION);
            dateReleveCell.setCellValue(bankLine.year() + "-" + bankLine.mounth());
            dateReleveCell.setCellStyle(cellStyle);

            Cell accountReleveCell = row.createCell(ID_ACOUNT_NUMBER_OF_RECONCILIATION);
            accountReleveCell.setCellValue(Double.parseDouble(bankLine.account().account()));
            accountReleveCell.setCellStyle(cellStyle);

            Cell labelAccountReleveCell = row.createCell(ID_ACOUNT_LABEL_OF_RECONCILIATION);
            labelAccountReleveCell.setCellValue(bankLine.account().label());
            labelAccountReleveCell.setCellStyle(cellStyle);

            Cell operationDateReleveCell = row.createCell(ID_OPERATION_DATE_OF_RECONCILIATION);
            operationDateReleveCell.setCellValue(bankLine.operationDate().format(DATE_FORMATTER));
            operationDateReleveCell.setCellStyle(cellStyle);

            Cell valueDateReleveCell = row.createCell(ID_VALUE_DATE_OF_RECONCILIATION);
            valueDateReleveCell.setCellValue(bankLine.valueDate().format(DATE_FORMATTER));
            valueDateReleveCell.setCellStyle(cellStyle);

            Cell labelReleveCell = row.createCell(ID_LABEL_OF_RECONCILIATION);
            labelReleveCell.setCellValue(bankLine.label());
            labelReleveCell.setCellStyle(cellStyle);

            Cell debitReleveCell = row.createCell(ID_DEBIT_OF_RECONCILIATION);
            debitReleveCell.setCellValue(bankLine.debit());
            debitReleveCell.setCellStyle(cellStyleAmount);

            Cell creditReleveCell = row.createCell(ID_CREDIT_OF_RECONCILIATION);
            creditReleveCell.setCellValue(bankLine.credit());
            creditReleveCell.setCellStyle(cellStyleAmount);
        }
        Cell messageCell = row.createCell(ID_COMMENT_OF_RECONCILIATION);
        messageCell.setCellValue(message);
        messageCell.setCellStyle(cellStyle);
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
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre);
            }
        }
        if (verifCell.getStringCellValue().equals(KO)) {
            verifCell.setCellStyle(getCellStyleVerifRed(style));
        } else {
            verifCell.setCellStyle(style);
        }

        Cell messageCell = row.createCell(ID_COMMENT_OF_LEDGER);
        if (link.getAddress() != null) {
            messageCell.setHyperlink(link);
        }
        messageCell.setCellValue(message.trim());
        addlineBlue(style, messageCell);
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

    private void addLabelCell(Line grandLivre, Row row, CellStyle style) {
        Cell labelCell = row.createCell(ID_LABEL_OF_LEDGER);
        if (!grandLivre.label().isEmpty()) {
            labelCell.setCellValue(grandLivre.label());
        } else {
            LOGGER.error("Le libellé est absent sur la ligne : {}", grandLivre);
        }
        labelCell.setCellStyle(style);
    }

    private void addCheckNumberCell(Line grandLivre, Row row, CellStyle style) {
        Cell checkNumberCell = row.createCell(ID_CHECK_OF_LEDGER);
        if (isDouble(grandLivre.checkNumber())) {
            checkNumberCell.setCellValue(Double.parseDouble(grandLivre.checkNumber()));
        } else {
            checkNumberCell.setCellValue(grandLivre.checkNumber());
        }
        checkNumberCell.setCellStyle(style);
    }

    private void addCounterPartCell(Line grandLivre, Row row, CellStyle style) {
        Cell counterPartCell = row.createCell(ID_COUNTERPART_NUMBER_OF_LEDGER);
        if (grandLivre.accountCounterpart() != null) {
            if (isDouble(grandLivre.accountCounterpart().account())) {
                counterPartCell.setCellValue(Double.parseDouble(grandLivre.accountCounterpart().account()));
            } else {
                counterPartCell.setCellValue(grandLivre.accountCounterpart().account());
            }
            Cell labelCounterPartCell = row.createCell(ID_COUNTERPART_LABEL_OF_LEDGER);
            labelCounterPartCell.setCellValue(grandLivre.accountCounterpart().label());
            labelCounterPartCell.setCellStyle(style);
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("La contre partie est absente sur la ligne : {}", grandLivre);
        }
        counterPartCell.setCellStyle(style);

    }

    private void addJournalCell(Line grandLivre, Row row, CellStyle style) {
        Cell journalCell = row.createCell(ID_JOURNAL_OF_LEDGER);
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
        journalCell.setCellStyle(style);
    }

    private void addDateCell(Line grandLivre, Row row, CellStyle style) {
        Cell dateCell = row.createCell(ID_DATE_OF_LEDGER);
        dateCell.setCellValue(grandLivre.date());
        dateCell.setCellStyle(style);
    }

    private void addDocumentCell(Line grandLivre, Row row, CellStyle cellStyle) {
        Cell documentCell = row.createCell(ID_DOCUMENT_OF_LEDGER);
        documentCell.setCellValue(grandLivre.document());
        if (isDouble(grandLivre.document())) {
            double document = Double.parseDouble(grandLivre.document());
            documentCell.setCellValue(document);
        } else if ((!grandLivre.label().contains(REPORT_DE)) && grandLivre.document().isEmpty()) {
            LOGGER.error("Le numéro de piéce est absente sur la ligne : {}", grandLivre);
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("Le numéro de piéce n'est pas numérique ({}) sur la ligne : {}", grandLivre.document(), grandLivre);
        }
        documentCell.setCellStyle(cellStyle);
    }

    private void addAccountCell(Line grandLivre, Row row, CellStyle cellStyle) {
        String numAccount = grandLivre.account().account();
        boolean isNotEmptyAccount = !numAccount.isEmpty();
        if (isNotEmptyAccount) {
            Cell accountNumberCell = row.createCell(ID_ACOUNT_NUMBER_OF_LEDGER);
            if (isDouble(numAccount)) {
                accountNumberCell.setCellValue(Double.parseDouble(numAccount));
            } else {
                accountNumberCell.setCellValue(numAccount);
            }
            accountNumberCell.setCellStyle(cellStyle);

            Cell accountLabelCell = row.createCell(ID_ACOUNT_LABEL_OF_LEDGER);
            accountLabelCell.setCellValue(grandLivre.account().label());
            accountLabelCell.setCellStyle(cellStyle);
        } else {
            LOGGER.error("Il manque le numéro de compte sur cette ligne : {}", grandLivre);
        }
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

    public void getCellsEnteteGrandLivre(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_GRAND_LIVRE) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(styleHeader);
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

    public void getCellsEnteteEtatRapprochement(Sheet sheet) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        CellStyle styleHeader = sheet.getWorkbook().createCellStyle();
        for (String label : NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(getCellStyleEntete(styleHeader));
        }
    }

    private CellStyle getCellStyleTotalAmount(Workbook workbook) {
        CellStyle style = getCellStyleAmount(workbook);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private CellStyle getCellStyleAmount(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(AMOUNT_FORMATTER));
        return style;
    }

    private CellStyle getCellStyleTotal(Workbook workbook) {
        CellStyle styleTotal = workbook.createCellStyle();
        styleTotal.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleTotal;
    }

    private CellStyle getCellStyleEntete(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    public void getLineOfExpenseKey(LineOfExpenseKey line, Row row) {
        Cell cell;
        for (int index = 0; index < NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES.length; index++) {
            cell = row.createCell(index);
            if (index == ID_LABEL_OF_LIST_OF_EXPENSES) {
                cell.setCellValue(line.label() + " : " + line.key() + " " + line.value());
            }
            cell.setCellStyle(getCellStyleTotal(row.getSheet().getWorkbook()));
        }
    }

    public void getLineOfExpenseTotal(LineOfExpenseTotal line, Row row) {
        Cell cell;
        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(getCellStyleTotal(row.getSheet().getWorkbook()));
        cell = row.createCell(ID_DATE_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(getCellStyleTotal(row.getSheet().getWorkbook()));
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
            cell.setCellStyle(getCellStyleTotal(row.getSheet().getWorkbook()));
        }
        if (!line.amount().isEmpty()) {
            cell = row.createCell(ID_AMOUNT_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.amount()));
            cell.setCellStyle(getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        if (!line.deduction().isEmpty()) {
            cell = row.createCell(ID_DEDUCTION_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.deduction()));
            cell.setCellStyle(getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        if (!line.recovery().isEmpty()) {
            cell = row.createCell(ID_RECOVERY_OF_LIST_OF_EXPENSES);
            cell.setCellValue(Double.parseDouble(line.recovery()));
            cell.setCellStyle(getCellStyleTotalAmount(row.getSheet().getWorkbook()));
        }
        cell = row.createCell(ID_COMMENT_OF_LIST_OF_EXPENSES);
        cell.setCellStyle(getCellStyleTotal(row.getSheet().getWorkbook()));
    }

    public void getLineOfExpense(LineOfExpense line, Row row, String pathDirectoryInvoice) {
        CreationHelper createHelper = row.getSheet().getWorkbook().getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);

        CellStyle styleColor;
        CellStyle styleAmountColor;
        if (row.getRowNum() % 2 == 0) {
            styleColor = getCellStyleWhite(row.getSheet().getWorkbook());
            styleAmountColor = getCellStyleAmountWhite(row.getSheet().getWorkbook());
        } else {
            styleColor = getCellStyleBlue(row.getSheet().getWorkbook());
            styleAmountColor = getCellStyleAmountBlue(row.getSheet().getWorkbook());
        }

        Cell cell;

        cell = row.createCell(ID_DOCUMENT_OF_LIST_OF_EXPENSES);
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

    private CellStyle getCellStyleWhite(Workbook workbook) {
        CellStyle styleWhite = workbook.createCellStyle();
        styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        styleWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleWhite;
    }

    private CellStyle getCellStyleBlue(Workbook workbook) {
        CellStyle styleBlue = workbook.createCellStyle();
        styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        styleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleBlue;
    }

    private CellStyle getCellStyleAmountWhite(Workbook workbook) {
        CellStyle styleAmountWhite = getCellStyleAmount(workbook);
        styleAmountWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        styleAmountWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleAmountWhite;
    }

    private CellStyle getCellStyleAmountBlue(Workbook workbook) {
        CellStyle styleAmountBlue = getCellStyleAmount(workbook);
        styleAmountBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        styleAmountBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleAmountBlue;
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

    public void writeDocumentMission(Workbook workbook, int idCellComment, int idCellDocumment) {
        Sheet sheetDocument = workbook.createSheet("Pieces manquante");
        TreeMap<String, String> ligneOfDocumentMissing = getListOfDocumentMissing(sheetDocument, idCellComment, idCellDocumment);
        writeListOfDocumentMissing(ligneOfDocumentMissing, sheetDocument);
    }
}