package life.light;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Map;

public class WriteFile {

    private static final Logger LOGGER = LogManager.getLogger();

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
            cell = headerRow.createCell(colNum++);
            cell.setCellValue("Libelle");
            rowNum++;
            for (Map.Entry<String, Account> entry : accounts.entrySet()) {
                Row row = sheet.createRow(rowNum);
                colNum = 0;
                row.createCell(colNum++).setCellValue(entry.getKey());
                row.createCell(colNum++).setCellValue(entry.getValue().label());
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

    public static void writeFileExcelGrandLivre(Object[] grandLivres, String printDate, String syndicName, String stopDate) {
        String exitFile = "." + File.separator + "temp" + File.separator
                + printDate.substring(6) + "-" + printDate.substring(3, 5) + "-" + printDate.substring(0, 2)
                + " Grand livre " + syndicName.substring(0, syndicName.length() - 1).trim()
                + " au " + stopDate.substring(6) + "-" + stopDate.substring(3, 5) + "-" + stopDate.substring(0, 2)
                + ".xlsx";
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();
            // Créer une nouvelle feuille dans le classeur
            Sheet sheet = workbook.createSheet("Grand Livre");
            int rowNum = 0;
            int cellNumEntete = 0;

            // Créer la ligne d'en-tête
            Row headerRow = sheet.createRow(rowNum);
            Cell cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Compte");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Intitulé du compte");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Pièce");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Date");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Journal");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Contrepartie");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("N° chèque");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Libellé");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Débit");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Crédit");
            cell.setCellStyle(getCellStyleEntete(workbook));
            cell = headerRow.createCell(cellNumEntete++);
            cell.setCellValue("Solde");
            cell.setCellStyle(getCellStyleEntete(workbook));
            rowNum++;

            for (Object grandLivre : grandLivres) {
                Row row = sheet.createRow(rowNum);
                int cellNum = 0;
                if (grandLivre instanceof Line) {
                    if (!((Line) grandLivre).account().account().isEmpty()) {
                        cell = row.createCell(cellNum);
                        cell.setCellValue(((Line) grandLivre).account().account());
                        if (isDouble(((Line) grandLivre).account().account())) {
                            double account = Double.parseDouble(((Line) grandLivre).account().account());
                            cell.setCellValue(account);
                        }
                        cell.setCellStyle(getCellStyleAlignmentLeft(workbook));
                    }
                    cellNum++;
                    if (!((Line) grandLivre).account().account().isEmpty()) {
                        row.createCell(cellNum).setCellValue(((Line) grandLivre).account().label());
                        row.getCell(cellNum).setCellStyle(getCellStyleAlignmentLeft(workbook));
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).document().isEmpty()) {
                        int document = Integer.parseInt(((Line) grandLivre).document());
                        row.createCell(cellNum).setCellValue(document);
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).date().isEmpty()) {
                        row.createCell(cellNum).setCellValue(((Line) grandLivre).date());
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).journal().isEmpty()) {
                        if (isDouble(((Line) grandLivre).journal())) {
                            int journal = Integer.parseInt(((Line) grandLivre).journal());
                            row.createCell(cellNum).setCellValue(journal);
                        } else {
                            row.createCell(cellNum).setCellValue(((Line) grandLivre).journal());
                        }
                        row.getCell(cellNum).setCellStyle(getCellStyleAlignmentLeft(workbook));
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                        row.getCell(cellNum).setCellStyle(getCellStyleAlignmentLeft(workbook));
                    }
                    cellNum++;
                    if (!((Line) grandLivre).counterpart().isEmpty()) {
                        if (isDouble(((Line) grandLivre).counterpart())) {
                            int counterpart = Integer.parseInt(((Line) grandLivre).counterpart());
                            row.createCell(cellNum).setCellValue(counterpart);
                        } else {
                            row.createCell(cellNum).setCellValue(((Line) grandLivre).counterpart());
                        }
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).checkNumber().isEmpty()) {
                        if (isDouble(((Line) grandLivre).checkNumber())) {
                            int checkNumber = Integer.parseInt(((Line) grandLivre).checkNumber());
                            row.createCell(cellNum).setCellValue(checkNumber);
                            row.getCell(cellNum).setCellStyle(getCellStyleAlignmentLeft(workbook));
                        } else {
                            row.createCell(cellNum).setCellValue(((Line) grandLivre).checkNumber());
                        }
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).label().isEmpty()) {
                        if (isDouble(((Line) grandLivre).checkNumber())) {
                            int label = Integer.parseInt(((Line) grandLivre).label());
                            row.createCell(cellNum).setCellValue(label);
                        } else {
                            row.createCell(cellNum).setCellValue(((Line) grandLivre).label());
                        }
                    } else {
                        row.createCell(cellNum).setCellValue(" ");
                    }
                    cellNum++;
                    if (!((Line) grandLivre).debit().isEmpty()) {
                        cell = row.createCell(cellNum);
                        cell.setCellValue(((Line) grandLivre).debit());
                        if (isDouble(((Line) grandLivre).debit())) {
                            double debit = Double.parseDouble(((Line) grandLivre).debit());
                            cell.setCellValue(debit);
                        }
                        cell.setCellStyle(getCellStyleAmount(workbook));
                    }
                    cellNum++;
                    if (!((Line) grandLivre).credit().isEmpty()) {
                        cell = row.createCell(cellNum);
                        cell.setCellValue(((Line) grandLivre).credit());
                        if (isDouble(((Line) grandLivre).credit())) {
                            double credit = Double.parseDouble(((Line) grandLivre).credit());
                            cell.setCellValue(credit);
                        }
                        cell.setCellStyle(getCellStyleAmount(workbook));
                    }
                }
                if (grandLivre instanceof TotalAccount) {
                    if (!((TotalAccount) grandLivre).account().account().isEmpty()) {
                        cell = row.createCell(0);
                        cell.setCellValue(((TotalAccount) grandLivre).account().account());
                        if (isDouble(((TotalAccount) grandLivre).account().account())) {
                            double account = Double.parseDouble(((TotalAccount) grandLivre).account().account());
                            cell.setCellValue(account);
                        }
                        cell.setCellStyle(getCellStyleTotal(workbook));
                    }
                    if (!((TotalAccount) grandLivre).account().label().isEmpty()) {
                        cell = row.createCell(1);
                        cell.setCellValue(((TotalAccount) grandLivre).account().label());
                        cell.setCellStyle(getCellStyleTotal(workbook));
                    }
                    if (!((TotalAccount) grandLivre).label().isEmpty()) {
                        cell = row.createCell(7);
                        cell.setCellValue(((TotalAccount) grandLivre).label());
                        cell.setCellStyle(getCellStyleTotal(workbook));
                    }
                    if (!((TotalAccount) grandLivre).debit().isEmpty()) {
                        cell = row.createCell(8);
                        cell.setCellValue(((TotalAccount) grandLivre).debit());
                        if (isDouble(((TotalAccount) grandLivre).debit())) {
                            double debit = Double.parseDouble(((TotalAccount) grandLivre).debit());
                            cell.setCellValue(debit);
                        }
                        cell.setCellStyle(getCellStyleTotalAmount(workbook));
                    }
                    if (!((TotalAccount) grandLivre).credit().isEmpty()) {
                        cell = row.createCell(9);
                        cell.setCellValue(((TotalAccount) grandLivre).credit());
                        if (isDouble(((TotalAccount) grandLivre).credit())) {
                            double credit = Double.parseDouble(((TotalAccount) grandLivre).credit());
                            cell.setCellValue(credit);
                        }
                        cell.setCellStyle(getCellStyleTotalAmount(workbook));
                    }
                }
                rowNum++;
            }
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

    private static CellStyle getCellStyleTotalAmount(Workbook workbook) {
        CellStyle styleTotalAmount = getCellStyleAmount(workbook);
        styleTotalAmount.setFont(getFontBold(workbook));
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