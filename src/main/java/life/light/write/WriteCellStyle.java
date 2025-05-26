package life.light.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class WriteCellStyle {

    public static final String AMOUNT_FORMATTER = "# ### ##0.00 €;[red]# ### ##0.00 €";

    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_WHITE = new XSSFColor(new java.awt.Color(255, 255, 255), null);
    private static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    private static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);

    CellStyle getCellStyleTotalAmount(Workbook workbook) {
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

    CellStyle getCellStyleTotal(Workbook workbook) {
        CellStyle styleTotal = workbook.createCellStyle();
        styleTotal.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleTotal;
    }

    CellStyle getCellStyleEntete(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    CellStyle getCellStyleWhite(Workbook workbook) {
        CellStyle styleWhite = workbook.createCellStyle();
        styleWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        styleWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleWhite;
    }

    CellStyle getCellStyleBlue(Workbook workbook) {
        CellStyle styleBlue = workbook.createCellStyle();
        styleBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        styleBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleBlue;
    }

    CellStyle getCellStyleAmountWhite(Workbook workbook) {
        CellStyle styleAmountWhite = getCellStyleAmount(workbook);
        styleAmountWhite.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        styleAmountWhite.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleAmountWhite;
    }

    CellStyle getCellStyleAmountBlue(Workbook workbook) {
        CellStyle styleAmountBlue = getCellStyleAmount(workbook);
        styleAmountBlue.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        styleAmountBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleAmountBlue;
    }

    CellStyle getCellStyleVerifRed(CellStyle style) {
        style.setFillForegroundColor(BACKGROUND_COLOR_RED);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
