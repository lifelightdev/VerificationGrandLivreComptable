package life.light.write;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class WriteCellStyle {

    public static final String AMOUNT_FORMATTER = "# ### ##0.00 €;[red]# ### ##0.00 €";

    public static final XSSFColor BACKGROUND_COLOR_BLUE = new XSSFColor(new java.awt.Color(240, 255, 255), null);
    public static final XSSFColor BACKGROUND_COLOR_WHITE = new XSSFColor(new java.awt.Color(255, 255, 255), null);
    private static final XSSFColor BACKGROUND_COLOR_GRAY = new XSSFColor(new java.awt.Color(200, 200, 200), null);
    private static final XSSFColor BACKGROUND_COLOR_RED = new XSSFColor(new java.awt.Color(255, 0, 0), null);
    private static final short fontSize = 12;
    private static final short fontSizeTotal = 14;

    public CellStyle getCellStyleTotalAmount(Workbook workbook) {
        CellStyle style = getCellStyleAmount(workbook);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(getFont(workbook, fontSizeTotal));
        return style;
    }

    private CellStyle getCellStyleAmount(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat(AMOUNT_FORMATTER));
        style.setFont(getFont(workbook, fontSize));
        return style;
    }

    private Font getFont(Workbook workbook, short fontSize) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints(fontSize);
        return font;
    }

    public CellStyle getCellStyleTotal(Workbook workbook) {
        CellStyle styleTotal = workbook.createCellStyle();
        styleTotal.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        styleTotal.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleTotal.setFont(getFont(workbook, fontSizeTotal));
        return styleTotal;
    }

    public CellStyle getCellStyleHeader(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(getFont(workbook, fontSize));
        return style;
    }

    public CellStyle getCellStyle(Workbook workbook, boolean isWhite) {
        CellStyle style = workbook.createCellStyle();
        if (isWhite) {
            style.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        } else {
            style.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        }
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(getFont(workbook, fontSize));
        return style;
    }

    public CellStyle getCellStyleAmount(Workbook workbook, boolean isWhite) {
        CellStyle styleAmount = getCellStyleAmount(workbook);
        if (isWhite) {
            styleAmount.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        } else {
            styleAmount.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        }
        styleAmount.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleAmount.setFont(getFont(workbook, fontSize));
        return styleAmount;
    }

    public CellStyle getCellStyleVerifRed(CellStyle style) {
        style.setFillForegroundColor(BACKGROUND_COLOR_RED);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
