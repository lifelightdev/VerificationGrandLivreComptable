package life.light.write;

import life.light.type.Line;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.List;

import static life.light.write.WriteOutil.*;

public class WriteCell {

    private static final Logger LOGGER = LogManager.getLogger();

    private WriteOutil writeOutil = new WriteOutil();

    public Cell addCellBalanceOfTotalInLedger(Row row, Cell debitCell, Cell creditCell, CellStyle styleTotalAmount) {
        Cell soldeCell = row.createCell(ID_BALANCE_OF_LEDGER);
        soldeCell.setCellFormula(debitCell.getAddress() + "-" + creditCell.getAddress());
        soldeCell.setCellStyle(styleTotalAmount);
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        return soldeCell;
    }

    public Cell addCellAmountOfTotalBuildingInLedger(Row row, int idDebitOfLedger, List<Integer> lineTotals, CellStyle styleTotalAmount) {
        Cell debitCell = row.createCell(idDebitOfLedger);
        StringBuilder sumDebit = new StringBuilder();
        for (Integer numRow : lineTotals) {
            sumDebit.append(CellReference.convertNumToColString(debitCell.getColumnIndex())).append(numRow).append("+");
        }
        debitCell.setCellFormula(sumDebit.substring(0, sumDebit.lastIndexOf("+")));
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(styleTotalAmount);
        return debitCell;
    }

    public void addCell(Row row, int idColum, String value, CellStyle style, String line, String name, String place) {
        Cell cell = row.createCell(idColum);
        WriteOutil writeOutil = new WriteOutil();
        if (writeOutil.isDouble(value)) {
            cell.setCellValue(Double.parseDouble(value));
        } else if (!value.isEmpty()) {
            cell.setCellValue(value);
        } else {
            if (name != null && !name.isEmpty()) {
                LOGGER.error("Il manque {} dans {} à la ligne {}", name, place, line);
            }
        }
        cell.setCellStyle(style);
    }

    public void addCellEmpty(int idFirstColum, int idLastColum, Row row, CellStyle style) {
        for (int idCell = idFirstColum; idCell < idLastColum; idCell++) {
            Cell cell = row.createCell(idCell);
            cell.setCellStyle(style);
        }
    }

    public Cell addCellCreditOfTotalAccountInLedger(Row row, int lastRowNumTotal, CellStyle cellStyleAmount) {
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        creditCell.setCellStyle(cellStyleAmount);
        CellAddress creditCellAddressFirst = new CellAddress(lastRowNumTotal + 1, creditCell.getAddress().getColumn());
        CellAddress creditCellAddressEnd = new CellAddress(creditCell.getAddress().getRow() - 1, creditCell.getAddress().getColumn());
        creditCell.setCellFormula("SUM(" + creditCellAddressFirst + ":" + creditCellAddressEnd + ")");
        return creditCell;
    }

    public Cell addCellDebitOfTotalAccountInLedger(Row row, int lastRowNumTotal, CellStyle cellStyleAmount) {
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        CellAddress debitCellAddressFirst = new CellAddress(lastRowNumTotal + 1, debitCell.getAddress().getColumn());
        CellAddress debitCellAddressEnd = new CellAddress(debitCell.getAddress().getRow() - 1, debitCell.getAddress().getColumn());
        debitCell.setCellFormula("SUM(" + debitCellAddressFirst + ":" + debitCellAddressEnd + ")");
        row.getSheet().getWorkbook().setForceFormulaRecalculation(true);
        debitCell.setCellStyle(cellStyleAmount);
        return debitCell;
    }

    public void addSoldeCell(Line grandLivre, Row row, CellStyle style, Cell creditCell, Cell debitCell) {
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

    public Cell addCreditCell(Line grandLivre, Row row, CellStyle style) {
        Cell creditCell = row.createCell(ID_CREDIT_OF_LEDGER);
        creditCell.setCellValue(grandLivre.credit());
        if (writeOutil.isDouble(grandLivre.credit())) {
            creditCell.setCellValue(Double.parseDouble(grandLivre.credit()));
        } else {
            creditCell.setCellValue(0);
        }
        creditCell.setCellStyle(style);
        return creditCell;
    }

    public Cell addDebitCell(Line grandLivre, Row row, CellStyle style) {
        Cell debitCell = row.createCell(ID_DEBIT_OF_LEDGER);
        debitCell.setCellValue(grandLivre.debit());
        if (writeOutil.isDouble(grandLivre.debit())) {
            debitCell.setCellValue(Double.parseDouble(grandLivre.debit()));
        } else {
            debitCell.setCellValue(0D);
        }
        debitCell.setCellStyle(style);
        return debitCell;
    }
}
