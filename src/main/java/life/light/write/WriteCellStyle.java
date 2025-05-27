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

    CellStyle getCellStyleHeader(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(BACKGROUND_COLOR_GRAY);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    CellStyle getCellStyle(Workbook workbook, boolean isWhite) {
        CellStyle style = workbook.createCellStyle();
        if (isWhite) {
            style.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        } else {
            style.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        }
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    CellStyle getCellStyleAmount(Workbook workbook, boolean isWhite) {
        CellStyle styleAmount = getCellStyleAmount(workbook);
        if (isWhite) {
            styleAmount.setFillForegroundColor(BACKGROUND_COLOR_WHITE);
        } else {
            styleAmount.setFillForegroundColor(BACKGROUND_COLOR_BLUE);
        }
        styleAmount.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return styleAmount;
    }

    CellStyle getCellStyleVerifRed(CellStyle style) {
        style.setFillForegroundColor(BACKGROUND_COLOR_RED);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
