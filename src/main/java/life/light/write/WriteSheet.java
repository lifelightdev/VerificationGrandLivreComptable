package life.light.write;

import life.light.type.Line;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;

import static life.light.write.WriteOutil.IMPOSSIBLE_DE_TROUVER_LA_PIECE;
import static life.light.write.WriteOutil.NOM_ENTETE_COLONNE_GRAND_LIVRE;

public class WriteSheet {
    private final WriteLine writeLine = new WriteLine();
    private final WriteOutil writeOutil = new WriteOutil();

    public void writeJournals(Object[] grandLivres, TreeSet<String> journals, String pathDirectoryInvoice, Workbook workbook) {
        int rowNum;
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
            writeLine.getCellsEnteteGrandLivre(sheetJournal);
            rowNum = 1;
            writeLine.getCellsEnteteGrandLivre(sheetJournal);
            for (Map.Entry<String, Line> line : ligneOfJournal.entrySet()) {
                Row row = sheetJournal.createRow(rowNum);
                writeLine.getLineGrandLivre(line.getValue(), row, false, pathDirectoryInvoice);
                rowNum++;
            }
            writeOutil.autoSizeCollum(NOM_ENTETE_COLONNE_GRAND_LIVRE.length, sheetJournal);
        }
    }

    public void writeDocumentMission(Workbook workbook, Sheet sheet, int idCellComment, int idCellDocumment) {
        Sheet sheetDocument = workbook.createSheet("Pieces manquante");
        TreeMap<String, String> ligneOfDocumentMissing = getListOfDocumentMissing(sheet, idCellComment, idCellDocumment);
        writeLine.getListOfDocumentMissing(ligneOfDocumentMissing, sheetDocument);
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
}
