package com.passargecorp.excel.workbook.output;

import org.apache.poi.ss.usermodel.Cell;

public class StringOutput implements OutputInterface {

    private final String stringValue;

    StringOutput(final String stringValue) {
        this.stringValue = stringValue;
    }

    public String getValue() {
        return stringValue;
    }

    public void setValue(Cell cell) {
        cell.setCellValue(getValue());
    }
}
