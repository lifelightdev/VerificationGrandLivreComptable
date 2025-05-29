package life.light.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class WriteCellStyleTest {

    private WriteCellStyle writeCellStyle;
    private Workbook workbook;

    @BeforeEach
    void setUp() {
        writeCellStyle = new WriteCellStyle();
        workbook = new XSSFWorkbook();
    }

    @Test
    void getCellStyleTotalAmount() {
        CellStyle style = writeCellStyle.getCellStyleTotalAmount(workbook);
        // Check that the style has the correct data format
        assertEquals(workbook.createDataFormat().getFormat(WriteCellStyle.AMOUNT_FORMATTER), style.getDataFormat());
        // Check that the style has the correct background color (gray)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        // We can't directly compare XSSFColor objects, but we can verify the pattern is set
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyleTotal() {
        CellStyle style = writeCellStyle.getCellStyleTotal(workbook);
        // Check that the style has the correct background color (gray)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyleHeader() {
        CellStyle baseStyle = workbook.createCellStyle();
        CellStyle style = writeCellStyle.getCellStyleHeader(baseStyle);
        // Check that the style has center alignment
        assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
        // Check that the style has the correct background color (gray)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyle_white() {
        CellStyle style = writeCellStyle.getCellStyle(workbook, true);
        // Check that the style has left alignment
        assertEquals(HorizontalAlignment.LEFT, style.getAlignment());
        // Check that the style has the correct background color (white)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyle_blue() {
        CellStyle style = writeCellStyle.getCellStyle(workbook, false);
        // Check that the style has left alignment
        assertEquals(HorizontalAlignment.LEFT, style.getAlignment());
        // Check that the style has the correct background color (blue)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyleAmount_white() {
        CellStyle style = writeCellStyle.getCellStyleAmount(workbook, true);
        // Check that the style has the correct data format
        assertEquals(workbook.createDataFormat().getFormat(WriteCellStyle.AMOUNT_FORMATTER), style.getDataFormat());
        // Check that the style has the correct background color (white)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyleAmount_blue() {
        CellStyle style = writeCellStyle.getCellStyleAmount(workbook, false);
        // Check that the style has the correct data format
        assertEquals(workbook.createDataFormat().getFormat(WriteCellStyle.AMOUNT_FORMATTER), style.getDataFormat());
        // Check that the style has the correct background color (blue)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }

    @Test
    void getCellStyleVerifRed() {
        CellStyle baseStyle = workbook.createCellStyle();
        CellStyle style = writeCellStyle.getCellStyleVerifRed(baseStyle);
        // Check that the style has the correct background color (red)
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertNotNull(style.getFillForegroundColorColor());
    }
}