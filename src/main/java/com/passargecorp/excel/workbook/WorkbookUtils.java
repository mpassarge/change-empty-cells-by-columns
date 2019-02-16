package com.passargecorp.excel.workbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class WorkbookUtils {

    public static String cellValueToString(final Cell cell){
        final CellType type = cell.getCellType();
        switch (type) {
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            case BLANK:
                return "";
            case _NONE:
                return "";
            default: throw new IllegalArgumentException();
        }
    }
}
