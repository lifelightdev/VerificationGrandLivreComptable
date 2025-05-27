package life.light.write;

import life.light.type.Line;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;

public class WriteOutil {

    private final WriteCellStyle  writeCellStyle = new WriteCellStyle();

    private static final Logger LOGGER = LogManager.getLogger();

    public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    public static final String[] NOM_ENTETE_COLONNE_LISTE_DES_DEPENSES = {"Pièce", "Date", "Libellé", "Montant",
            "Déduction", "Récuperation", "Commentaire"};
    public static final int ID_DOCUMENT_OF_LIST_OF_EXPENSES = 0;
    public static final int ID_DATE_OF_LIST_OF_EXPENSES = 1;
    public static final int ID_LABEL_OF_LIST_OF_EXPENSES = 2;
    public static final int ID_AMOUNT_OF_LIST_OF_EXPENSES = 3;
    public static final int ID_DEDUCTION_OF_LIST_OF_EXPENSES = 4;
    public static final int ID_RECOVERY_OF_LIST_OF_EXPENSES = 5;
    public static final int ID_COMMENT_OF_LIST_OF_EXPENSES = 6;

    protected static final String[] NOM_ENTETE_COLONNE_GRAND_LIVRE = {"Compte", "Intitulé du compte", "Pièce", "Date",
            "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "Solde (Calculé)", "Vérification", "Commentaire"};
    public static final int ID_ACOUNT_NUMBER_OF_LEDGER = 0;
    public static final int ID_ACOUNT_LABEL_OF_LEDGER = 1;
    public static final int ID_DOCUMENT_OF_LEDGER = 2;
    public static final int ID_DATE_OF_LEDGER = 3;
    public static final int ID_JOURNAL_OF_LEDGER = 4;
    public static final int ID_COUNTERPART_NUMBER_OF_LEDGER = 5;
    public static final int ID_COUNTERPART_LABEL_OF_LEDGER = 6;
    public static final int ID_CHECK_OF_LEDGER = 7;
    public static final int ID_LABEL_OF_LEDGER = 8;
    public static final int ID_DEBIT_OF_LEDGER = 9;
    public static final int ID_CREDIT_OF_LEDGER = 10;
    public static final int ID_BALANCE_OF_LEDGER = 11;
    public static final int ID_VERIFFICATION_OF_LEDGER = 12;
    public static final int ID_COMMENT_OF_LEDGER = 13;

    protected static final String[] NOM_ENTETE_COLONNE_ETAT_RAPPROCHEMENT = {"Compte", "Intitulé du compte", "Pièce",
            "Date", "Journal", "Contrepartie", "Intitulé de la contrepartie", "N° chèque", "Libellé", "Débit", "Crédit",
            "----",
            "Mois du relevé", "Compte", "Intitulé du compte", "Date de l'opération", "Date de valeur", "Libellé",
            "Débit", "Crédit", "Commentaire"};
    public static final int ID_MONTH_OF_SATEMENT_OF_RECONCILIATION = 12;
    public static final int ID_ACOUNT_NUMBER_OF_RECONCILIATION = 13;
    public static final int ID_ACOUNT_LABEL_OF_RECONCILIATION = 14;
    public static final int ID_OPERATION_DATE_OF_RECONCILIATION = 15;
    public static final int ID_VALUE_DATE_OF_RECONCILIATION = 16;
    public static final int ID_LABEL_OF_RECONCILIATION = 17;
    public static final int ID_DEBIT_OF_RECONCILIATION = 18;
    public static final int ID_CREDIT_OF_RECONCILIATION = 19;
    public static final int ID_COMMENT_OF_RECONCILIATION = 20;

    public static final String REPORT_DE = "Report de";
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

    public void addVerifCells(Line grandLivre, Row row, CellStyle style, String pathDirectoryInvoice) {
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
                LOGGER.info("{} sur cette ligne : {}", message, grandLivre.toString());
            }
        }
        if (verifCell.getStringCellValue().equals(KO)) {
            verifCell.setCellStyle(writeCellStyle.getCellStyleVerifRed(style));
        } else {
            verifCell.setCellStyle(style);
        }

        Cell messageCell = row.createCell(ID_COMMENT_OF_LEDGER);
        if (link.getAddress() != null) {
            messageCell.setHyperlink(link);
        }
        messageCell.setCellValue(message.trim());
        messageCell.setCellStyle(style);
    }

    public String getMessageFindDocument(String document, Hyperlink link, String thePathDirectoryInvoice) {
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
                verifCell.setCellStyle(style);
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

    public void autoSizeCollum(int numberOfColumns, Sheet sheet) {
        for (int idCollum = 0; idCollum < numberOfColumns; idCollum++) {
            sheet.autoSizeColumn(idCollum);
        }
        sheet.createFreezePane(0, 1);
    }


}
