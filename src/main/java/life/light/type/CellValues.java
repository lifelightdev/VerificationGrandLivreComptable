package life.light.type;

import org.apache.poi.ss.usermodel.CellStyle;

public record CellValues(int idColum, String value, CellStyle style, String line, String name) {
}
