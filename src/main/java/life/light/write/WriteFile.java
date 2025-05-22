package life.light.write;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;

import static life.light.write.OutilWrite.*;

public class WriteFile {

    private static final Logger LOGGER = LogManager.getLogger();
    public static final String POINTAGE_RELEVE_OK = "Pointage Relevé OK";
    public static final String POINTAGE_GL_OK = "Pointage GL OK";
    private OutilWrite outilWrite = new OutilWrite();

    // TODO faire la gestion des fichiers (existe, n'existe pas, pas de dossier ...)

    public void writeFileCSVAccounts(Map<String, TypeAccount> accounts, String fileName) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(fileName))) {
            TreeMap<String, TypeAccount> map = new TreeMap<>(accounts);
            writer.write("Compte;Intitulé du compte;" + System.lineSeparator());
            for (Map.Entry<String, TypeAccount> accountEntry : map.entrySet()) {
                String line = accountEntry.getValue().account() + " ; " +
                        accountEntry.getValue().label() + " ; " +
                        System.lineSeparator();
                writer.write(line);
            }
            LOGGER.info("L'écriture du fichier {} est terminée.", fileName);
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public void writeFileExcelAccounts(Map<String, TypeAccount> map, String fileName) {
        TreeMap<String, TypeAccount> accounts = new TreeMap<>(map);
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            // Style
            CellStyle styleWhite = workbook.createCellStyle();
            styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
            styleWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            styleWhite.setAlignment(HorizontalAlignment.LEFT);
            CellStyle styleBlue = workbook.createCellStyle();
            styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
            styleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            styleBlue.setAlignment(HorizontalAlignment.LEFT);

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
            cell.setCellValue("Intitulé du compte");

            rowNum++;
            for (Map.Entry<String, TypeAccount> entry : accounts.entrySet()) {
                Row row = sheet.createRow(rowNum);
                Cell accountNumberCell = outilWrite.getAccountNumberCell(entry.getKey(), row, numColAccount);
                Cell labelCell = row.createCell(numColLabelle);
                labelCell.setCellValue(entry.getValue().label());
                if (rowNum % 2 == 0) {
                    accountNumberCell.setCellStyle(styleWhite);
                    labelCell.setCellStyle(styleWhite);
                } else {
                    accountNumberCell.setCellStyle(styleBlue);
                    labelCell.setCellStyle(styleBlue);
                }
                rowNum++;
            }
            sheet.createFreezePane(0, 1);
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            // Écrire le contenu du classeur dans un fichier
            outilWrite.writeWorkbook(fileName, workbook);
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public void writeFileCSVGrandLivre(Object[] grandLivres) {
        String exitFile = ".\\temp\\GrandLivre.csv";
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(exitFile))) {
            StringBuilder line = new StringBuilder("Pièce; Date; Compte; Journal; Contrepartie; N° chèque; Libellé; Débit; Crédit;");
            line.append(System.lineSeparator());
            writer.write(line.toString());
            for (Object grandLivre : grandLivres) {
                line = new StringBuilder();
                if (grandLivre instanceof Line) {
                    line.append(((Line) grandLivre).document()).append(" ; ");
                    line.append(((Line) grandLivre).date()).append(" ; ");
                    if (((Line) grandLivre).account() != null) {
                        line.append(((Line) grandLivre).account().account()).append(" ; ");
                    } else {
                        line.append(" ; ");
                    }
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
                    if (((TotalAccount) grandLivre).account() != null) {
                        line.append(((TotalAccount) grandLivre).account().account()).append(" ; ");
                    } else {
                        line.append(" ; ");
                    }
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

    public void writeFileExcelGrandLivre(Object[] grandLivres, String nameFile, TreeSet<String> journals, String pathDirectoryInvoice) {
        String exitFile = "." + File.separator + "temp" + File.separator + nameFile;
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            // Style
            DataFormat dataFormat = workbook.createDataFormat();
            Short dataAmount = dataFormat.getFormat("# ### ##0.00 €;[red]# ### ##0.00 €");
            CellStyle styleTotal = outilWrite.getCellStyleTotal(workbook.createCellStyle());
            CellStyle styleTotalAmount = outilWrite.getCellStyleTotalAmount(workbook.createCellStyle(), dataAmount);
            CellStyle styleWhite = workbook.createCellStyle();
            styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
            styleWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CellStyle styleBlue = workbook.createCellStyle();
            styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
            styleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            CellStyle styleAmountWhite = outilWrite.getCellStyleAmount(workbook.createCellStyle(), dataAmount);
            styleAmountWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
            CellStyle styleAmountBlue = outilWrite.getCellStyleAmount(workbook.createCellStyle(), dataAmount);
            styleAmountBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
            CellStyle styleHeader = workbook.createCellStyle();

            // Créer une nouvelle feuille dans le classeur pour le grand livre
            Sheet sheet = workbook.createSheet("Grand Livre");
            outilWrite.getCellsEnteteGrandLivre(sheet, styleHeader);
            int rowNum = 1;
            int lastRowNumTotal = 0;
            List<Integer> lineTotals = new ArrayList<>();
            for (Object grandLivre : grandLivres) {
                Row row = sheet.createRow(rowNum);
                if (grandLivre instanceof Line) {
                    if (rowNum % 2 == 0) {
                        outilWrite.getLineGrandLivre((Line) grandLivre, row, styleBlue, styleAmountBlue, workbook, true, pathDirectoryInvoice);
                    } else {
                        outilWrite.getLineGrandLivre((Line) grandLivre, row, styleWhite, styleAmountWhite, workbook, true, pathDirectoryInvoice);
                    }
                }
                if (grandLivre instanceof TotalAccount) {
                    outilWrite.getTotalAccount((TotalAccount) grandLivre, row, styleTotal, styleTotalAmount, workbook, lastRowNumTotal);
                    lastRowNumTotal = rowNum;
                    lineTotals.add(rowNum + 1);
                }
                if (grandLivre instanceof TotalBuilding) {
                    outilWrite.getTotalBuilding((TotalBuilding) grandLivre, row, styleTotal, styleTotalAmount, workbook, lineTotals);
                }
                rowNum++;
            }
            int cellNumEntete = NOM_ENTETE_COLONNE_GRAND_LIVRE.length;
            for (int idCollum = 0; idCollum < cellNumEntete; idCollum++) {
                sheet.autoSizeColumn(idCollum);
            }
            sheet.createFreezePane(0, 1);

            // Créer une nouvelle feuille par journal
            for (String journal : journals) {
                Sheet sheetJournal = workbook.createSheet(journal);
                TreeMap<String, Line> ligneOfJournal = new TreeMap<>();
                for (Object grandLivre : grandLivres) {
                    if (grandLivre instanceof Line) {
                        if (journal.equals(((Line) grandLivre).journal())) {
                            ligneOfJournal.put(((Line) grandLivre).document(), (Line) grandLivre);
                        }
                    }
                }
                outilWrite.getCellsEnteteGrandLivre(sheetJournal, styleHeader);
                rowNum = 1;
                outilWrite.getCellsEnteteGrandLivre(sheetJournal, styleHeader);
                for (Map.Entry<String, Line> line : ligneOfJournal.entrySet()) {
                    Row row = sheetJournal.createRow(rowNum);
                    if (rowNum % 2 == 0) {
                        outilWrite.getLineGrandLivre(line.getValue(), row, styleBlue, styleAmountBlue, workbook, false, pathDirectoryInvoice);
                    } else {
                        outilWrite.getLineGrandLivre(line.getValue(), row, styleWhite, styleAmountWhite, workbook, false, pathDirectoryInvoice);
                    }
                    rowNum++;
                }
                for (int idCollum = 0; idCollum < cellNumEntete; idCollum++) {
                    sheetJournal.autoSizeColumn(idCollum);
                }
                sheetJournal.createFreezePane(0, 1);
            }

            // Écrire le contenu du classeur dans un fichier
            outilWrite.writeWorkbook(exitFile, workbook);
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    public void writeFileExcelEtatRaprochement(List<Line> grandLivres, String exitFile, List<BankLine> bankLines) {
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();
            // Style
            DataFormat dataFormat = workbook.createDataFormat();
            Short dataAmount = dataFormat.getFormat("# ### ##0.00 €;[red]# ### ##0.00 €");
            CellStyle styleWhite = workbook.createCellStyle();
            styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
            CellStyle styleBlue = workbook.createCellStyle();
            styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
            CellStyle styleAmountWhite = outilWrite.getCellStyleAmount(workbook.createCellStyle(), dataAmount);
            styleAmountWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
            CellStyle styleAmountBlue = outilWrite.getCellStyleAmount(workbook.createCellStyle(), dataAmount);
            styleAmountBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
            CellStyle styleHeader = workbook.createCellStyle();

            pointageReleve(new ArrayList<>(grandLivres), new ArrayList<>(bankLines), workbook,
                    styleHeader, styleBlue, styleAmountBlue, styleWhite, styleAmountWhite);
            pointageGL(new ArrayList<>(grandLivres), new ArrayList<>(bankLines), workbook,
                    styleHeader, styleBlue, styleAmountBlue, styleWhite, styleAmountWhite);

            if (workbook.getSheet(POINTAGE_RELEVE_OK).getLastRowNum() != workbook.getSheet(POINTAGE_GL_OK).getLastRowNum()) {
                LOGGER.error(("Il n'y a pas le même nombre d'opération pointées OK dans le grand livre et sur les relevés de compte"));
            }

            // Écrire le contenu du classeur dans un fichier
            outilWrite.writeWorkbook(exitFile, workbook);
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    private void pointageReleve(List<Line> grandLivres, List<BankLine> bankLines, Workbook workbook,
                                CellStyle styleHeader, CellStyle styleBlue, CellStyle styleAmountBlue,
                                CellStyle styleWhite, CellStyle styleAmountWhite) {
        Sheet sheetPointage = workbook.createSheet(POINTAGE_RELEVE_OK);
        List<Line> grandLivresKO = new ArrayList<>();
        List<BankLine> bankLinesKO = new ArrayList<>();
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage, styleHeader);
        int rowNumPointage = 1;
        for (Line grandLivre : grandLivres) {
            Row row = sheetPointage.createRow(rowNumPointage);
            BankLine bankLineFound = null;
            String message = KO;
            for (BankLine bankLine : bankLines) {
                // TODO il faut verifier la date aussi
                if (grandLivre.account().account().equals(bankLine.account().account())) {
                    if (!grandLivre.credit().isEmpty() && Double.parseDouble(grandLivre.credit()) == bankLine.debit()) {
                        message = OK;
                        bankLineFound = bankLine;
                        break;
                    }
                    if (!grandLivre.debit().isEmpty() && Double.parseDouble(grandLivre.debit()) == bankLine.credit()) {
                        message = OK;
                        bankLineFound = bankLine;
                        break;
                    }
                    bankLinesKO.add(bankLine);
                }
            }
            if (message.equals(KO)) {
                grandLivresKO.add(grandLivre);
            } else {
                message = "Correspondante entre le grand livre et les relevés de banque";
                if (rowNumPointage % 2 == 0) {
                    outilWrite.getLineEtatRapprochement(grandLivre, row, styleBlue, styleAmountBlue, bankLineFound, message);
                } else {
                    outilWrite.getLineEtatRapprochement(grandLivre, row, styleWhite, styleAmountWhite, bankLineFound, message);
                }
                bankLines.remove(bankLineFound);
                rowNumPointage++;
            }
        }
        for (int idCollum = 0; idCollum < NOM_ENTETE_COLONNE_GRAND_LIVRE.length; idCollum++) {
            sheetPointage.autoSizeColumn(idCollum);
        }
        sheetPointage.createFreezePane(0, 1);

        Sheet sheetPointage1 = workbook.createSheet("Pointage Relevé KO");
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage1, styleHeader);
        rowNumPointage = 1;
        for (Line grandLivre : grandLivresKO) {
            Row row = sheetPointage1.createRow(rowNumPointage);
            BankLine bankLineFound = null;
            String message = KO;
            for (BankLine bankLine : bankLinesKO) {
                // TODO il faut verifier la date aussi
                if (!grandLivre.credit().isEmpty() && Double.parseDouble(grandLivre.credit()) == bankLine.debit()) {
                    message = OK;
                    bankLineFound = bankLine;
                    break;
                }
                if (!grandLivre.debit().isEmpty() && Double.parseDouble(grandLivre.debit()) == bankLine.credit()) {
                    message = OK;
                    bankLineFound = bankLine;
                    break;
                }
            }
            if (message.equals(KO)) {
                message = "Aucune correspondance du grand livre dans les relevés de banque";
                if (rowNumPointage % 2 == 0) {
                    outilWrite.getLineEtatRapprochement(grandLivre, row, styleBlue, styleAmountBlue, bankLineFound, message);
                } else {
                    outilWrite.getLineEtatRapprochement(grandLivre, row, styleWhite, styleAmountWhite, bankLineFound, message);
                }
                bankLines.remove(bankLineFound);
                rowNumPointage++;
            }
        }
        for (int idCollum = 0; idCollum < NOM_ENTETE_COLONNE_GRAND_LIVRE.length; idCollum++) {
            sheetPointage1.autoSizeColumn(idCollum);
        }
        sheetPointage1.createFreezePane(0, 1);
    }

    private void pointageGL(List<Line> grandLivres, List<BankLine> bankLines, Workbook workbook,
                            CellStyle styleHeader, CellStyle styleBlue, CellStyle styleAmountBlue,
                            CellStyle styleWhite, CellStyle styleAmountWhite) {
        Sheet sheetPointage = workbook.createSheet(POINTAGE_GL_OK);
        List<Line> grandLivresKO = new ArrayList<>();
        List<BankLine> bankLinesKO = new ArrayList<>();
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage, styleHeader);
        int rowNumPointage = 1;
        for (BankLine bankLine : bankLines) {
            Row row = sheetPointage.createRow(rowNumPointage);
            Line lineGLFound = null;
            String message = KO;
            for (Line grandLivre : grandLivres) {
                // TODO il faut verifier la date aussi
                if (grandLivre.account().account().equals(bankLine.account().account())) {
                    if (!grandLivre.credit().isEmpty() && Double.parseDouble(grandLivre.credit()) == bankLine.debit()) {
                        message = OK;
                        lineGLFound = grandLivre;
                        break;
                    }
                    if (!grandLivre.debit().isEmpty() && Double.parseDouble(grandLivre.debit()) == bankLine.credit()) {
                        message = OK;
                        lineGLFound = grandLivre;
                        break;
                    }
                    grandLivresKO.add(grandLivre);
                }
            }
            if (message.equals(KO)) {
                bankLinesKO.add(bankLine);
            } else {
                message = "Correspondante entre le grand livre et les relevés de banque";
                if (rowNumPointage % 2 == 0) {
                    outilWrite.getLineEtatRapprochement(lineGLFound, row, styleBlue, styleAmountBlue, bankLine, message);
                } else {
                    outilWrite.getLineEtatRapprochement(lineGLFound, row, styleWhite, styleAmountWhite, bankLine, message);
                }
                grandLivres.remove(lineGLFound);
                rowNumPointage++;
            }
        }
        for (int idCollum = 0; idCollum < NOM_ENTETE_COLONNE_GRAND_LIVRE.length; idCollum++) {
            sheetPointage.autoSizeColumn(idCollum);
        }
        sheetPointage.createFreezePane(0, 1);

        Sheet sheetPointage1 = workbook.createSheet("Pointage GL KO");
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage1, styleHeader);
        rowNumPointage = 1;
        for (BankLine bankLine : bankLinesKO) {
            Row row = sheetPointage1.createRow(rowNumPointage);
            Line lineGrandLivreFound = null;
            String message = KO;
            for (Line grandLivre : grandLivresKO) {
                // TODO il faut verifier la date aussi
                if (!grandLivre.credit().isEmpty() && Double.parseDouble(grandLivre.credit()) == bankLine.debit()) {
                    message = OK;
                    lineGrandLivreFound = grandLivre;
                    break;
                }
                if (!grandLivre.debit().isEmpty() && Double.parseDouble(grandLivre.debit()) == bankLine.credit()) {
                    message = OK;
                    lineGrandLivreFound = grandLivre;
                    break;
                }
            }
            if (message.equals(KO)) {
                message = "Aucune correspondance du relevé de banque dans le grand livre";

                if (rowNumPointage % 2 == 0) {
                    outilWrite.getLineEtatRapprochement(lineGrandLivreFound, row, styleBlue, styleAmountBlue, bankLine, message);
                } else {
                    outilWrite.getLineEtatRapprochement(lineGrandLivreFound, row, styleWhite, styleAmountWhite, bankLine, message);
                }
                grandLivres.remove(lineGrandLivreFound);
                rowNumPointage++;
            }
        }
        for (int idCollum = 0; idCollum < NOM_ENTETE_COLONNE_GRAND_LIVRE.length; idCollum++) {
            sheetPointage1.autoSizeColumn(idCollum);
        }
        sheetPointage1.createFreezePane(0, 1);
    }

}