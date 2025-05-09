package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class WriteFile {

    // RGB
    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    public static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);
    private static final Logger LOGGER = LogManager.getLogger();
    private static final String[] NOM_ENTETE_COLONNE = {"Compte", "Intitulé du compte", "Pièce", "Date", "Journal",
            "Contrepartie", "N° chèque", "Libellé", "Débit", "Crédit", "Solde (Calculé)", "Vérification", "Commentaire"};
    public static final String REPORT_DE = "Report de";
    public static final String PATH_DIRECTORY_INVOICE = "";
    public static final String CLASSE_6 = "6";
    public static final String KO = "KO";

    // TODO faire la gestion des fichiers (existe, n'existe pas, pas de dossier ...)

    public static void writeFileCSVAccounts(Map<String, Account> accounts, String fileName) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(fileName))) {
            TreeMap<String, Account> map = new TreeMap<>(accounts);
            StringBuilder lineFirst = new StringBuilder("Compte;Intitulé du compte;");
            lineFirst.append(System.lineSeparator());
            writer.write(lineFirst.toString());
            for (Map.Entry<String, Account> accountEntry : map.entrySet()) {
                StringBuilder line = new StringBuilder();
                line.append(accountEntry.getValue().account()).append(" ; ");
                line.append(accountEntry.getValue().label()).append(" ; ");
                line.append(System.lineSeparator());
                writer.write(line.toString());
            }
            LOGGER.info("L'écriture du fichier {} est terminée.", fileName);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public static void writeFileExcelAccounts(Map<String, Account> map, String fileName) {
        TreeMap<String, Account> accounts = new TreeMap<>(map);
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();
            // Créer une nouvelle feuille dans le classeur
            Sheet sheet = workbook.createSheet("Plan comptable");
            int rowNum = 0;
            // Créer la ligne d'en-tête
            Row headerRow = sheet.createRow(rowNum);
            int numColAccount = 0;
            Cell cell = headerRow.createCell(numColAccount);
            cell.setCellValue("Compte");
            int numColLabelle = 1;
            cell = headerRow.createCell(numColLabelle);
            cell.setCellValue("Libelle");

            rowNum++;
            for (Map.Entry<String, Account> entry : accounts.entrySet()) {
                Row row = sheet.createRow(rowNum);
                row.createCell(numColAccount).setCellValue(entry.getKey());
                row.createCell(numColLabelle).setCellValue(entry.getValue().label());
                rowNum++;
            }
            sheet.createFreezePane(0, 1);
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            // Écrire le contenu du classeur dans un fichier
            try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
                workbook.write(outputStream);
                LOGGER.info("L'écriture du fichier {} est terminée.", fileName);
            } catch (IOException e) {
                LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
            }
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public static void writeFileGrandLivre(Object[] grandLivres) {
        String exitFile = ".\\temp\\GrandLivre.csv";
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(exitFile))) {
            StringBuilder line = new StringBuilder("Pièce; Date; Compte; Journal; Contrepartie; N° chèque; Libellé; Débit; Crédit;");
            line.append(System.lineSeparator());
            writer.write(line.toString());
            for (Object grandLivre : grandLivres) {
                if (grandLivre instanceof Line) {
                    line.append(((Line) grandLivre).document()).append(" ; ");
                    line.append(((Line) grandLivre).date()).append(" ; ");
                    line.append(((Line) grandLivre).account().account()).append(" ; ");
                    line.append(((Line) grandLivre).journal()).append(" ; ");
                    if (((Line) grandLivre).accountCounterpart() != null) {
                        line.append(((Line) grandLivre).accountCounterpart().account()).append(" ; ");
                    } else {
                        line.append(" ; ");
                    }
                    line.append(((Line) grandLivre).checkNumber()).append(" ; ");
                    line.append(((Line) grandLivre).label()).append(" ; ");
                    line.append(((Line) grandLivre).debit()).append(" ; ");
                    line.append(((Line) grandLivre).credit()).append(" ; ");
                    line.append(System.lineSeparator());
                }
                if (grandLivre instanceof TotalAccount) {
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(((TotalAccount) grandLivre).account().account()).append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(((TotalAccount) grandLivre).label()).append(" ; ");
                    line.append(((TotalAccount) grandLivre).debit()).append(" ; ");
                    line.append(((TotalAccount) grandLivre).credit()).append(" ; ");
                    line.append(System.lineSeparator());
                }
                if (grandLivre instanceof TotalBuilding) {
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(" ; ");
                    line.append(((TotalBuilding) grandLivre).label()).append(" ; ");
                    line.append(((TotalBuilding) grandLivre).debit()).append(" ; ");
                    line.append(((TotalBuilding) grandLivre).credit()).append(" ; ");
                    line.append(System.lineSeparator());
                }
                writer.write(line.toString());
            }
            LOGGER.info("L'écriture du fichier {} est terminée.", exitFile);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    public static void writeFileExcelGrandLivre(Object[] grandLivres, String nameFile) {
        String exitFile = "." + File.separator + "temp" + File.separator + nameFile;
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();
            // Créer une nouvelle feuille dans le classeur
            Sheet sheet = workbook.createSheet("Grand Livre");
            getCellsEntete(sheet, workbook);
            int rowNum = 1;
            int lastRowNumTotal = 0;
            List<Integer> lineTotals = new ArrayList<>();
            for (Object grandLivre : grandLivres) {
                Row row = sheet.createRow(rowNum);
                if (grandLivre instanceof Line) {
                    getLine((Line) grandLivre, row, workbook);
                }
                if (grandLivre instanceof TotalAccount) {
                    getTotalAccount((TotalAccount) grandLivre, row, workbook, lastRowNumTotal);
                    lastRowNumTotal = rowNum;
                    lineTotals.add(rowNum + 1);
                }
                if (grandLivre instanceof TotalBuilding) {
                    getTotalBuilding((TotalBuilding) grandLivre, row, workbook, lineTotals);
                }
                rowNum++;
            }
            int cellNumEntete = NOM_ENTETE_COLONNE.length;
            for (int idCollum = 0; idCollum < cellNumEntete; idCollum++) {
                sheet.autoSizeColumn(idCollum);
            }
            sheet.createFreezePane(0, 1);
            // Écrire le contenu du classeur dans un fichier
            try (FileOutputStream outputStream = new FileOutputStream(exitFile)) {
                workbook.write(outputStream);
                LOGGER.info("L'écriture du fichier {} est terminée.", exitFile);
            } catch (IOException e) {
                LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
            }
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    private static void getTotalBuilding(TotalBuilding grandLivre, Row row, Workbook workbook, List<Integer> lineTotals) {
        Cell cell;
        for (int idCell = 0; idCell < 8; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        int cellNum = 7;
        if (!grandLivre.label().isEmpty()) {
            cell = row.createCell(cellNum++);
            cell.setCellValue(grandLivre.label());
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        Cell debitCell = row.createCell(cellNum++);
        StringBuilder sumDebit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumDebit.append(CellReference.convertNumToColString(debitCell.getColumnIndex())).append(numRow).append("+");
        }
        debitCell.setCellFormula(sumDebit.substring(0, sumDebit.lastIndexOf("+")));
        workbook.setForceFormulaRecalculation(true);
        debitCell.setCellStyle(getCellStyleTotalAmount(workbook));

        Cell creditCell = row.createCell(cellNum++);
        StringBuilder sumCredit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumCredit.append(CellReference.convertNumToColString(creditCell.getColumnIndex())).append(numRow).append("+");
        }
        creditCell.setCellFormula(sumCredit.substring(0, sumCredit.lastIndexOf("+")));
        workbook.setForceFormulaRecalculation(true);
        creditCell.setCellStyle(getCellStyleTotalAmount(workbook));

        Cell soldeCell = row.createCell(cellNum);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(getCellStyleTotalAmount(workbook));
        workbook.setForceFormulaRecalculation(true);
    }

    private static void getTotalAccount(TotalAccount grandLivre, Row row, Workbook workbook, int lastRowNumTotal) {
        Cell cell;
        if (!grandLivre.account().account().isEmpty()) {
            cell = row.createCell(0);
            cell.setCellValue(grandLivre.account().account());
            if (isDouble(grandLivre.account().account())) {
                double account = Double.parseDouble(grandLivre.account().account());
                cell.setCellValue(account);
            }
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        if (!grandLivre.account().label().isEmpty()) {
            cell = row.createCell(1);
            cell.setCellValue(grandLivre.account().label());
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        for (int idCell = 2; idCell < 7; idCell++) {
            cell = row.createCell(idCell);
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        int cellNum = 7;
        if (!grandLivre.label().isEmpty()) {
            cell = row.createCell(cellNum++);
            cell.setCellValue(grandLivre.label());
            cell.setCellStyle(getCellStyleTotal(workbook));
        }
        Cell debitCell = row.createCell(cellNum++);
        CellAddress debitCellAddressFirst = new CellAddress(lastRowNumTotal + 1, debitCell.getAddress().getColumn());
        CellAddress debitCellAddressEnd = new CellAddress(debitCell.getAddress().getRow() - 1, debitCell.getAddress().getColumn());
        debitCell.setCellFormula("SUM(" + debitCellAddressFirst + ":" + debitCellAddressEnd + ")");
        workbook.setForceFormulaRecalculation(true);
        debitCell.setCellStyle(getCellStyleTotalAmount(workbook));

        Cell creditCell = row.createCell(cellNum++);
        creditCell.setCellStyle(getCellStyleTotalAmount(workbook));
        CellAddress creditCellAddressFirst = new CellAddress(lastRowNumTotal + 1, creditCell.getAddress().getColumn());
        CellAddress creditCellAddressEnd = new CellAddress(creditCell.getAddress().getRow() - 1, creditCell.getAddress().getColumn());
        creditCell.setCellFormula("SUM(" + creditCellAddressFirst + ":" + creditCellAddressEnd + ")");

        Cell soldeCell = row.createCell(cellNum++);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(getCellStyleTotalAmount(workbook));
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
                    cellVerif.setCellStyle(getCellStyleTotalAmount(workbook));
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
            cellVerif.setCellStyle(getCellStyleVerifRed(workbook));
            LOGGER.info("Vérification du total du compte [{}] Solde PDF [{}] Solde Excel [{}] débit PDF [{}] débit Excel [{}] crédit PDF [{}] crédit Excel [{}] ligne [{}]",
                    grandLivre.account().account(), amount, solde,
                    debitGrandLivre, debitExcel,
                    creditGrandLivre, creditExcel,
                    grandLivre);
        }
        cellVerif.setCellValue(verif);

        Cell cellMessqage = row.createCell(cellNum++);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount) + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(getCellStyleTotalAmount(workbook));
    }

    private static void getLine(Line grandLivre, Row row, Workbook workbook) {
        int cellNum = 0;

        cellNum = addAccountCell(grandLivre, row, workbook, cellNum);
        cellNum = addDocumentCell(grandLivre, row, workbook, cellNum);
        cellNum = addDateCell(grandLivre, row, workbook, cellNum);
        cellNum = addJournalCell(grandLivre, row, workbook, cellNum);
        cellNum = addCounterPartCell(grandLivre, row, workbook, cellNum);
        cellNum = addCheckNumberCell(grandLivre, row, workbook, cellNum);
        cellNum = addLabelCell(grandLivre, row, workbook, cellNum);

        Cell debitCell = addDebitCell(grandLivre, row, workbook, cellNum);
        cellNum++;

        Cell creditCell = addCreditCell(grandLivre, row, workbook, cellNum);
        cellNum++;

        cellNum = addSoldeCell(grandLivre, row, workbook, cellNum, creditCell, debitCell);

        cellNum = addVerifCells(grandLivre, row, workbook, cellNum);

    }

    private static int addVerifCells(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        String message = "";
        CreationHelper createHelper = workbook.getCreationHelper();
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
        Cell verifCell = row.createCell(cellNum);
        if (grandLivre.label().startsWith(REPORT_DE)) {
            message = getMessageVerifLineReport(grandLivre, row, workbook, verifCell, message);
        } else {
            if (grandLivre.account().account().startsWith(CLASSE_6)
                    || (grandLivre.account().account().startsWith("401")) && (grandLivre.accountCounterpart().account().startsWith(CLASSE_6))){
                message = getMessageFindDocument(grandLivre, verifCell, message, link);
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
            verifCell.setCellStyle(getCellStyleVerifRed(workbook));
        } else {
            addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), verifCell);
        }

        cellNum++;
        Cell messageCell = row.createCell(cellNum);
        if (link.getAddress() != null) {
            messageCell.setHyperlink(link);
        }
        messageCell.setCellValue(message.trim());
        addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), messageCell);
        return cellNum;
    }

    private static String getMessageFindDocument(Line grandLivre, Cell verifCell, String message, Hyperlink link) {
        File pathDirectoryInvoice = new File(PATH_DIRECTORY_INVOICE);
        File[] files = pathDirectoryInvoice.listFiles();
        if (null != files) {
            boolean find = false;
            for (File fichier : files) {
                if (fichier.isFile()) {
                    if (fichier.getName().contains(grandLivre.document())) {
                        verifCell.setCellValue("OK");
                        find = true;
                        message = fichier.getAbsoluteFile().toString().replace("F:","D:");
                        link.setAddress(fichier.toURI().toString().replace("F:","D:"));
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
                                    message = fichierDuSousDossier.getAbsoluteFile().toString().replace("F:","D:");
                                    link.setAddress(fichierDuSousDossier.toURI().toString().replace("F:","D:"));
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            if (!find) {
                message = "Impossible de trouver la pièce";
                LOGGER.info(message + " : " + grandLivre.document()
                        + " dans le dossier : " + pathDirectoryInvoice
                        + " sur le compte " + grandLivre.account().account()
                        + " libelle de l'opération " + grandLivre.label()
                );
            }
        }
        return message;
    }

    private static String getMessageVerifLineReport(Line grandLivre, Row row, Workbook workbook, Cell verifCell, String message) {
        CellStyle style = getCellStyleVerifRed(workbook);
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
                addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), verifCell);
            } else {
                message = "Le montant du report est de " + amount + " le solde est de  " + Double.parseDouble(nombreFormate)
                        + " le débit est de " + Double.parseDouble(grandLivre.debit()) + " le credit est de " + Double.parseDouble(grandLivre.credit());
                LOGGER.info(message + " sur le compte " + grandLivre.account().account());
                verifCell.setCellValue(KO);
            }
        }
        return message;
    }

    private static String getAmountInLineReport(Line grandLivre) {
        return grandLivre.label().substring(REPORT_DE.length(), grandLivre.label().length() - 1).trim().replace(" ", "");
    }

    private static int addSoldeCell(Line grandLivre, Row row, Workbook workbook, int cellNum, Cell creditCell, Cell debitCell) {
        Cell soldeCell = row.createCell(cellNum);
        String formule;
        if (grandLivre.label().startsWith(REPORT_DE)) {
            formule = creditCell.getAddress() + "-" + debitCell.getAddress();
        } else {
            int rowIndex = soldeCell.getRowIndex() - 1;
            int col = soldeCell.getColumnIndex();
            CellAddress beforeSoldeCellAddress = new CellAddress(rowIndex, col);
            formule = beforeSoldeCellAddress + "+" + debitCell.getAddress() + "-" + creditCell.getAddress();
        }
        soldeCell.setCellFormula(formule);
        addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), soldeCell);
        cellNum++;
        return cellNum;
    }

    private static Cell addCreditCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell creditCell = row.createCell(cellNum);
        creditCell.setCellValue(grandLivre.credit());
        if (isDouble(grandLivre.credit())) {
            creditCell.setCellValue(Double.parseDouble(grandLivre.credit()));
        } else {
            creditCell.setCellValue(0);
        }
        addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), creditCell);
        return creditCell;
    }

    private static Cell addDebitCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell debitCell = row.createCell(cellNum);
        debitCell.setCellValue(grandLivre.debit());
        if (isDouble(grandLivre.debit())) {
            debitCell.setCellValue(Double.parseDouble(grandLivre.debit()));
        } else {
            debitCell.setCellValue(0);
        }
        addlineBlue(row.getRowNum(), getCellStyleAmount(workbook), debitCell);
        return debitCell;
    }

    private static int addLabelCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell labelCell = row.createCell(cellNum);
        if (!grandLivre.label().isEmpty()) {
            labelCell.setCellValue(grandLivre.label());
        } else {
            LOGGER.error("Le libellé est absent sur la ligne : {}", grandLivre);
        }
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), labelCell);
        cellNum++;
        return cellNum;
    }

    private static int addCheckNumberCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell checkNumberCell = row.createCell(cellNum);
        checkNumberCell.setCellValue(grandLivre.checkNumber());
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), checkNumberCell);
        cellNum++;
        return cellNum;
    }

    private static int addCounterPartCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell counterPartCell = row.createCell(cellNum);
        if (grandLivre.accountCounterpart() != null) {
            counterPartCell.setCellValue(grandLivre.accountCounterpart().account());
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("La contre partie est absente sur la ligne : {}", grandLivre);
        }
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), counterPartCell);
        cellNum++;
        return cellNum;
    }

    private static int addJournalCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell journalCell = row.createCell(cellNum);
        journalCell.setCellValue(grandLivre.journal());
        if (!grandLivre.journal().isEmpty()) {
            journalCell.setCellValue(grandLivre.journal());
        } else if (!grandLivre.label().contains(REPORT_DE)) {
            LOGGER.error("Le journal est absent sur la ligne : {}", grandLivre);
        }
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), journalCell);
        cellNum++;
        return cellNum;
    }

    private static int addDateCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        Cell dateCell = row.createCell(cellNum);
        dateCell.setCellValue(grandLivre.date());
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), dateCell);
        cellNum++;
        return cellNum;
    }

    private static int addDocumentCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
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
        addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), documentCell);
        cellNum++;
        return cellNum;
    }

    private static int addAccountCell(Line grandLivre, Row row, Workbook workbook, int cellNum) {
        String numAccount = grandLivre.account().account();
        boolean isNotEmptyAccount = !numAccount.isEmpty();
        Cell accountNumberCell = row.createCell(cellNum++);
        accountNumberCell.setCellValue(numAccount);
        if (isNotEmptyAccount) {
            accountNumberCell.setCellValue(numAccount);
            addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), accountNumberCell);
            Cell accountLabelCell = row.createCell(cellNum++);
            accountLabelCell.setCellValue(grandLivre.account().label());
            addlineBlue(row.getRowNum(), getCellStyleAlignmentLeft(workbook), accountLabelCell);
        } else {
            LOGGER.error("Il manque le numéro de compte sur cette ligne : {}", grandLivre);
        }
        return cellNum;
    }

    private static CellStyle getCellStyleVerifRed(Workbook workbook) {
        CellStyle style = getCellStyleAlignmentLeft(workbook);
        style.setFont(getFontBold(workbook));
        style.setFillForegroundColor(BACKGROUND_COLOR_RED);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static void addlineBlue(int rowNum, CellStyle cellStyle, Cell cell) {
        if (rowNum % 2 == 0) {
            cellStyle = getCellStyleBlue(cellStyle);
        }
        cell.setCellStyle(cellStyle);
    }

    private static void getCellsEntete(Sheet sheet, Workbook workbook) {
        int index = 0;
        Row headerRow = sheet.createRow(index);
        for (String label : NOM_ENTETE_COLONNE) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(label);
            cell.setCellStyle(getCellStyleEntete(workbook));
        }
    }

    private static CellStyle getCellStyleBlue(CellStyle cellStyle) {
        cellStyle.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    private static CellStyle getCellStyleTotalAmount(Workbook workbook) {
        CellStyle styleTotalAmount = getCellStyleAmount(workbook);
        styleTotalAmount.setFont(getFontBold(workbook));
        styleTotalAmount.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotalAmount.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleTotalAmount;
    }

    private static CellStyle getCellStyleAmount(Workbook workbook) {
        CellStyle styleAmount = workbook.createCellStyle();
        styleAmount.setDataFormat(getFormatAmount(workbook));
        return styleAmount;
    }

    private static short getFormatAmount(Workbook workbook) {
        DataFormat dataFormat = workbook.createDataFormat();
        return dataFormat.getFormat("# ### ##0.00 €;[red]# ### ##0.00 €");
    }

    private static CellStyle getCellStyleTotal(Workbook workbook) {
        CellStyle styleTotal = getCellStyleAlignmentLeft(workbook);
        styleTotal.setFont(getFontBold(workbook));
        styleTotal.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleTotal;
    }

    private static CellStyle getCellStyleAlignmentLeft(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    private static CellStyle getCellStyleEntete(Workbook workbook) {
        Font font = getFontBold(workbook);
        CellStyle styleEntete = workbook.createCellStyle();
        styleEntete.setAlignment(HorizontalAlignment.CENTER);
        styleEntete.setFont(font);
        styleEntete.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleEntete.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleEntete;
    }

    private static Font getFontBold(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        return font;
    }

    private static boolean isDouble(String str) {
        if (str == null || str.isEmpty()) {
            return false; // Une chaîne nulle ou vide ne peut pas être un double
        }
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

}