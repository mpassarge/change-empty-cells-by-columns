package com.passargecorp.excel.workbook.output;

import org.apache.poi.ss.usermodel.Cell;

public class ZeroOutput implements OutputInterface {

    public Double getValue() {
        return 0.0;
    }

    public void setValue(Cell cell) {
        cell.setCellValue(getValue());
    }
}
