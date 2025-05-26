package life.light.write;

import life.light.type.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
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
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", fileName, e.getMessage());
        }
    }

    public void writeFileExcelAccounts(Map<String, TypeAccount> map, String fileName) {
        TreeMap<String, TypeAccount> accounts = new TreeMap<>(map);
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            // Créer une nouvelle feuille dans le classeur
            Sheet sheet = workbook.createSheet("Plan comptable");
            int rowNum = 0;
            // Créer la ligne d'en-tête
            Row headerRow = sheet.createRow(rowNum);
            Cell cell = headerRow.createCell(0);
            cell.setCellValue("Compte");
            cell = headerRow.createCell(1);
            cell.setCellValue("Intitulé du compte");

            rowNum++;
            for (Map.Entry<String, TypeAccount> entry : accounts.entrySet()) {
                Row row = sheet.createRow(rowNum);
                Cell cellAccountNumber = row.createCell(0);
                if (outilWrite.isDouble(entry.getKey())) {
                    cellAccountNumber.setCellValue(Double.parseDouble(entry.getKey()));
                } else {
                    cellAccountNumber.setCellValue(entry.getKey());
                }
                Cell labelCell = row.createCell(1);
                labelCell.setCellValue(entry.getValue().label());
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

    public void writeFileCSVGrandLivre(Object[] grandLivres, String exitFile) {

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
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", exitFile, e.getMessage());
        }
    }

    public void writeFileExcelGrandLivre(Object[] grandLivres, String pathNameFile, TreeSet<String> journals, String pathDirectoryInvoice) {
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            // Créer une nouvelle feuille dans le classeur pour le grand livre
            Sheet sheet = workbook.createSheet("Grand Livre");
            outilWrite.getCellsEnteteGrandLivre(sheet);
            int rowNum = 1;
            int lastRowNumTotal = 0;
            List<Integer> lineTotals = new ArrayList<>();
            for (Object grandLivre : grandLivres) {
                if (grandLivre != null) {
                    Row row = sheet.createRow(rowNum);
                    if (grandLivre instanceof Line) {
                        outilWrite.getLineGrandLivre((Line) grandLivre, row, true, pathDirectoryInvoice);
                    }
                    if (grandLivre instanceof TotalAccount) {
                        outilWrite.getTotalAccount((TotalAccount) grandLivre, row, lastRowNumTotal);
                        lastRowNumTotal = rowNum;
                        lineTotals.add(rowNum + 1);
                    }
                    if (grandLivre instanceof TotalBuilding) {
                        outilWrite.getTotalBuilding((TotalBuilding) grandLivre, row, lineTotals);
                    }
                    rowNum++;
                }
            }

            outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheet);

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
                outilWrite.getCellsEnteteGrandLivre(sheetJournal);
                rowNum = 1;
                outilWrite.getCellsEnteteGrandLivre(sheetJournal);
                for (Map.Entry<String, Line> line : ligneOfJournal.entrySet()) {
                    Row row = sheetJournal.createRow(rowNum);
                    outilWrite.getLineGrandLivre(line.getValue(), row, false, pathDirectoryInvoice);
                    rowNum++;
                }
                outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetJournal);
            }

            // Créer une nouvelle feuille pour les pieces manquantes
            outilWrite.writeDocumentMission(workbook, sheet, ID_COMMENT_OF_LEDGER, ID_DOCUMENT_OF_LEDGER);

            // Écrire le contenu du classeur dans un fichier
            outilWrite.writeWorkbook(pathNameFile, workbook);
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", pathNameFile, e.getMessage());
        }
    }

    public void writeFileExcelListeDesDepenses(Object[] listeDesDepenses, String pathNameFile, String pathDirectoryInvoice) {
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            // Créer une nouvelle feuille dans le classeur pour le grand livre
            Sheet sheet = workbook.createSheet("Liste des dépenses");
            outilWrite.getCellsEnteteListeDesDepenses(sheet);
            int rowNum = 1;
            for (Object line : listeDesDepenses) {
                if (line != null) {
                    Row row = sheet.createRow(rowNum);
                    if (line instanceof LineOfExpenseKey) {
                        outilWrite.getLineOfExpenseKey((LineOfExpenseKey) line, row);
                    }
                    if (line instanceof LineOfExpenseTotal) {
                        outilWrite.getLineOfExpenseTotal((LineOfExpenseTotal) line, row);
                    }
                    if (line instanceof LineOfExpense) {
                        outilWrite.getLineOfExpense((LineOfExpense) line, row, pathDirectoryInvoice);
                    }
                    rowNum++;
                }
            }
            outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES.length, sheet);

            // Créer une nouvelle feuille pour les pieces manquantes
            outilWrite.writeDocumentMission(workbook, sheet, ID_COMMENT_OF_LIST_OF_EXPENSES, ID_DOCUMENT_OF_LIST_OF_EXPENSES);

            // Écrire le contenu du classeur dans un fichier
            outilWrite.writeWorkbook(pathNameFile, workbook);
            // Fermer le classeur
            workbook.close();
        } catch (IOException e) {
            LOGGER.error("Erreur lors de l'écriture dans le fichier de sortie '{}': {}", pathNameFile, e.getMessage());
        }
    }

    public void writeFileExcelEtatRaprochement(List<Line> grandLivres, String exitFile, List<BankLine> bankLines) {
        try {
            // Créer un nouveau classeur Excel
            Workbook workbook = new XSSFWorkbook();

            pointageReleve(new ArrayList<>(grandLivres), new ArrayList<>(bankLines), workbook);
            pointageGL(new ArrayList<>(grandLivres), new ArrayList<>(bankLines), workbook);

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

    private void pointageReleve(List<Line> grandLivres, List<BankLine> bankLines, Workbook workbook) {
        Sheet sheetPointage = workbook.createSheet(POINTAGE_RELEVE_OK);
        List<Line> grandLivresKO = new ArrayList<>();
        List<BankLine> bankLinesKO = new ArrayList<>();
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage);
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
                outilWrite.getLineEtatRapprochement(grandLivre, row, bankLineFound, message);
                bankLines.remove(bankLineFound);
                rowNumPointage++;
            }
        }
        outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetPointage);

        Sheet sheetPointage1 = workbook.createSheet("Pointage Relevé KO");
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage1);
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
                outilWrite.getLineEtatRapprochement(grandLivre, row, bankLineFound, message);
                bankLines.remove(bankLineFound);
                rowNumPointage++;
            }
        }
        outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetPointage1);
    }

    private void pointageGL(List<Line> grandLivres, List<BankLine> bankLines, Workbook workbook) {
        Sheet sheetPointage = workbook.createSheet(POINTAGE_GL_OK);
        List<Line> grandLivresKO = new ArrayList<>();
        List<BankLine> bankLinesKO = new ArrayList<>();
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage);
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
                outilWrite.getLineEtatRapprochement(lineGLFound, row, bankLine, message);
                grandLivres.remove(lineGLFound);
                rowNumPointage++;
            }
        }
        outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetPointage);

        Sheet sheetPointage1 = workbook.createSheet("Pointage GL KO");
        outilWrite.getCellsEnteteEtatRapprochement(sheetPointage1);
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
                outilWrite.getLineEtatRapprochement(lineGrandLivreFound, row, bankLine, message);
                grandLivres.remove(lineGrandLivreFound);
                rowNumPointage++;
            }
        }
        outilWrite.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetPointage1);
    }
}