package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.Map;

public class WriteFile {

    // RGB
    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    public static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);
    private static final Logger LOGGER = LogManager.getLogger();
    private static final String[] NOM_ENTETE_COLONNE = {"Compte", "Intitulé du compte", "Pièce", "Date", "Journal",
            "Contrepartie", "N° chèque", "Libellé", "Débit", "Crédit", "Solde (Calculé)", "Vérification", "Commentaire"};

    // TODO faire la gestion des fichiers (existe, n'existe pas, pas de dossier ...)

    public static void writeFileAccounts(Map<String, Account> accounts) {
        String exitFile = "." + File.separator + "temp" + File.separator + "ListeDesCompte.csv";
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(exitFile))) {
            for (Map.Entry<String, Account> accountEntry : accounts.entrySet()) {
                String line = accountEntry.getValue().account() + "; " + accountEntry.getValue().label() + "\n";
                writer.write(line);
            }
            LOGGER.info("L'écriture du fichier {} est terminée.", exitFile);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    public static void writeFileExcelAccounts(Map<String, Account> accounts) {
        String exitFile = "." + File.separator + "temp" + File.separator + "Plan comptable.xlsx";
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();
            // Créer une nouvelle feuille dans le classeur
            Sheet sheet = workbook.createSheet("Plan comptable");
            int rowNum = 0;
            // Créer la ligne d'en-tête
            int colNum = 0;
            Row headerRow = sheet.createRow(colNum);
            Cell cell = headerRow.createCell(colNum++);
            cell.setCellValue("Compte");
            sheet.autoSizeColumn(colNum);
            cell = headerRow.createCell(colNum++);
            cell.setCellValue("Libelle");
            sheet.autoSizeColumn(colNum);
            rowNum++;
            for (Map.Entry<String, Account> entry : accounts.entrySet()) {
                Row row = sheet.createRow(rowNum);
                colNum = 0;
                row.createCell(colNum++).setCellValue(entry.getKey());
                row.createCell(colNum).setCellValue(entry.getValue().label());
                rowNum++;
            }
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

    public static void writeFileGrandLivre(Object[] grandLivres) {
        String exitFile = ".\\temp\\GrandLivre.csv";
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(exitFile))) {
            String line = "Pièce; Date; Compte; Journal; Contrepartie; N° chèque; Libellé; Débit; Crédit;\n";
            writer.write(line);
            for (Object grandLivre : grandLivres) {
                line = "";
                if (grandLivre instanceof Line) {
                    if (!((Line) grandLivre).document().isEmpty()) {
                        line += ((Line) grandLivre).document() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).date().isEmpty()) {
                        line += ((Line) grandLivre).date() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).account().account().isEmpty()) {
                        line += ((Line) grandLivre).account().account() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).journal().isEmpty()) {
                        line += ((Line) grandLivre).journal() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).counterpart().isEmpty()) {
                        line += ((Line) grandLivre).counterpart() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).checkNumber().isEmpty()) {
                        line += ((Line) grandLivre).checkNumber() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).label().isEmpty()) {
                        line += ((Line) grandLivre).label() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).debit().isEmpty()) {
                        line += ((Line) grandLivre).debit() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((Line) grandLivre).credit().isEmpty()) {
                        line += ((Line) grandLivre).credit() + "; ";
                    } else {
                        line += " ; ";
                    }
                    line += "\n";
                }
                if (grandLivre instanceof TotalAccount) {
                    if (!((TotalAccount) grandLivre).account().account().isEmpty()) {
                        line += ((TotalAccount) grandLivre).account().account() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((TotalAccount) grandLivre).label().isEmpty()) {
                        line += ((TotalAccount) grandLivre).label() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((TotalAccount) grandLivre).debit().isEmpty()) {
                        line += ((TotalAccount) grandLivre).debit() + "; ";
                    } else {
                        line += " ; ";
                    }
                    if (!((TotalAccount) grandLivre).credit().isEmpty()) {
                        line += ((TotalAccount) grandLivre).credit() + "; ";
                    } else {
                        line += " ; ";
                    }
                    line += "\n";
                }
                writer.write(line);
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
            for (Object grandLivre : grandLivres) {
                Row row = sheet.createRow(rowNum);
                if (grandLivre instanceof Line) {
                    getLine((Line) grandLivre, row, workbook, rowNum);
                }
                if (grandLivre instanceof TotalAccount) {
                    getTotalAccount((TotalAccount) grandLivre, row, workbook, lastRowNumTotal);
                    lastRowNumTotal = rowNum;
                }
                rowNum++;
            }
            int cellNumEntete = NOM_ENTETE_COLONNE.length;
            for (int idCollum = 0; idCollum < cellNumEntete; idCollum++) {
                sheet.autoSizeColumn(idCollum);
            }
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
        String verif = "";
        workbook.setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        double solde = (double) Math.round(evaluator.evaluate(soldeCell).getNumberValue() * 100) / 100;
        double debitExcel = (double) Math.round(evaluator.evaluate(debitCell).getNumberValue() * 100) / 100;
        Double debitGrandLivre = Double.parseDouble(grandLivre.debit());
        double creditExcel = (double) Math.round(evaluator.evaluate(creditCell).getNumberValue() * 100) / 100;
        Double creditGrandLivre = Double.parseDouble(grandLivre.credit());
        if (Double.parseDouble(amount) == solde) {
            if (debitExcel == debitGrandLivre) {
                if (creditExcel == creditGrandLivre) {
                    verif = "OK";
                    cellVerif.setCellStyle(getCellStyleTotalAmount(workbook));
                } else {
                    verif = "KO";
                }
            } else {
                verif = "KO";
            }
        } else {
            verif = "KO";
        }
        if (verif.equals("KO")) {
            cellVerif.setCellStyle(getCellStyleVerifRed(workbook));
            LOGGER.info("Vérification du total du compte [{}] Solde PDF [{}] Solde Excel [{}] débit PDF [{}] débit Excel [{}] crédit PDF [{}] crédit Excel [{}] ligne [{}]",
                    grandLivre.account().account(), amount, solde,
                    debitGrandLivre, debitExcel,
                    creditGrandLivre, creditExcel,
                    grandLivre);
        }
        cellVerif.setCellValue(verif);

        Cell cellMessqage = row.createCell(cellNum++);
        String formuleIfSolde = "IF(ROUND(" + soldeCell.getAddress() + ",2)=" + Double.parseDouble(amount)
                + ", \" \", \"Le solde n'est pas égale \")";
        String formuleIfCredit = "IF(" + creditCell.getAddress() + "=" + creditGrandLivre + ", " + formuleIfSolde + ", \"Le total credit n'est pas égale " + creditGrandLivre + " \")";
        String formuleIfDebit = "IF(" + debitCell.getAddress() + "=" + debitGrandLivre + ", " + formuleIfCredit + ", \"Le total débit n'est pas égale " + debitGrandLivre + " \")";

        cellMessqage.setCellFormula(formuleIfDebit);
        cellMessqage.setCellStyle(getCellStyleTotalAmount(workbook));
    }

    private static void getLine(Line grandLivre, Row row, Workbook workbook, int rowNum) {
        Cell cell;
        int cellNum = 0;
        String numAccount = grandLivre.account().account();
        boolean isNotEmptyAccount = !numAccount.isEmpty();
        double debit = 0;
        double credit = 0;
        cell = row.createCell(cellNum);
        cell.setCellValue(numAccount);
        if (isNotEmptyAccount) {
            if (isDouble(numAccount)) {
                double account = Double.parseDouble(numAccount);
                cell.setCellValue(account);
            }
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        if (isNotEmptyAccount) {
            cell.setCellValue(grandLivre.account().label());
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.document());
        if (isDouble(grandLivre.document())) {
            double document = Double.parseDouble(grandLivre.document());
            cell.setCellValue(document);
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.date());
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.journal());
        if (isDouble(grandLivre.journal())) {
            double journal = Double.parseDouble(grandLivre.journal());
            cell.setCellValue(journal);
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.counterpart());
        if (isDouble(grandLivre.counterpart())) {
            double counterpart = Double.parseDouble(grandLivre.counterpart());
            cell.setCellValue(counterpart);
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.checkNumber());
        if (isDouble(grandLivre.checkNumber())) {
            double checkNumber = Double.parseDouble(grandLivre.checkNumber());
            cell.setCellValue(checkNumber);
        }
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        cell = row.createCell(cellNum);
        cell.setCellValue(grandLivre.label());
        addlineBlue(rowNum, getCellStyleAlignmentLeft(workbook), cell);

        cellNum++;
        Cell debitCell = row.createCell(cellNum);
        debitCell.setCellValue(grandLivre.debit());
        if (isDouble(grandLivre.debit())) {
            debit = Double.parseDouble(grandLivre.debit());
            debitCell.setCellValue(debit);
        } else {
            debitCell.setCellValue(0);
        }
        addlineBlue(rowNum, getCellStyleAmount(workbook), debitCell);

        cellNum++;
        Cell creditCell = row.createCell(cellNum);
        creditCell.setCellValue(grandLivre.credit());
        if (isDouble(grandLivre.credit())) {
            credit = Double.parseDouble(grandLivre.credit());
            creditCell.setCellValue(credit);
        } else {
            creditCell.setCellValue(0);
        }
        addlineBlue(rowNum, getCellStyleAmount(workbook), creditCell);

        cellNum++;
        Cell soldeCell = row.createCell(cellNum);
        String formule;
        if (grandLivre.label().startsWith("Report de ")) {
            formule = creditCell.getAddress() + "-" + debitCell.getAddress();
        } else {
            int rowIndex = soldeCell.getRowIndex() - 1;
            int col = soldeCell.getColumnIndex();
            CellAddress beforeSoldeCellAddress = new CellAddress(rowIndex, col);
            formule = beforeSoldeCellAddress + "+" + debitCell.getAddress() + "-" + creditCell.getAddress();
        }
        soldeCell.setCellFormula(formule);
        addlineBlue(rowNum, getCellStyleAmount(workbook), soldeCell);

        String message = "";
        cellNum++;
        cell = row.createCell(cellNum);
        addlineBlue(rowNum, getCellStyleAmount(workbook), cell);
        if (grandLivre.label().startsWith("Report de ")) {
            CellStyle style = getCellStyleVerifRed(workbook);
            cell.setCellStyle(style);
            String amount = grandLivre.label().substring("Report de ".length(), grandLivre.label().length() - 1).trim().replace(" ", "");
            if (isDouble(amount)) {
                double amountDouble = Double.parseDouble(amount);
                DecimalFormat df = new DecimalFormat("#.00");
                String nombreFormate = df.format((debit - credit)).replace(",", ".");
                if (amountDouble == Double.parseDouble(nombreFormate)) {
                    cell.setCellValue("OK");
                    addlineBlue(rowNum, getCellStyleAmount(workbook), cell);
                } else {
                    message = "Le montant du report est de " + amount + " le solde est de  " + Double.parseDouble(nombreFormate)
                            + " le débit est de " + debit + " le credit est de " + credit;
                    LOGGER.info(message + " sur le compte " + grandLivre.account().account());
                    cell.setCellValue("KO");
                }
            }
        } else {
            if (grandLivre.account().account().startsWith("6")) {
                File pathDirectory = new File("D:\\Le Nidor\\2024\\FACTURES");
                File[] files = pathDirectory.listFiles();
                if (null != files) {
                    boolean find = false;
                    for (File fichier : files) {
                        if (fichier.isFile()) {
                            if (fichier.getName().contains(grandLivre.document())) {
                                cell.setCellValue("OK");
                                find = true;
                                break;
                            }
                            message = "Impossible de trouver la facture lier à la piece : " + grandLivre.document()
                                    + " dans le dossier : " + fichier.getAbsolutePath().replace(fichier.getName(), "");
                        } else if (fichier.isDirectory()) {
                            File[] sousDossier = fichier.listFiles();
                            if (null != sousDossier) {
                                boolean findSousDossier = false;
                                for (File fichierDuSousDossier : sousDossier) {
                                    if (fichierDuSousDossier.isFile()) {
                                        if (fichierDuSousDossier.getName().contains(grandLivre.document())) {
                                            cell.setCellValue("KO");
                                            findSousDossier = true;
                                            break;
                                        }
                                    }
                                }
                                if (!findSousDossier) {
                                    message = "Impossible de trouver la facture lier à la piece : " + grandLivre.document()
                                            + " dans le sous dossier : " + fichier.getPath();
                                    LOGGER.info(message + " sur le compte " + grandLivre.account().account() + " pour la ligne : " + grandLivre);
                                }
                            }
                        }
                    }
                    if (!find) {
                        message = "Impossible de trouver la facture lier à la piece : " + grandLivre.document()
                                + "dans le dossier : " + pathDirectory + "\n";
                        LOGGER.info(message + " sur le compte " + grandLivre.account().account());
                    }
                }
            }
        }

        if (cell.getStringCellValue().equals("KO")) {
            cellNum++;
            cell = row.createCell(cellNum);
            cell.setCellValue(message.trim());
            addlineBlue(rowNum, getCellStyleAmount(workbook), cell);
        }

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
        return dataFormat.getFormat("# ##0.00 €;[red]# ##0.00 €");
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