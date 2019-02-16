package com.passargecorp.excel.workbook.output;

import org.apache.poi.ss.usermodel.Cell;

public interface OutputInterface {

    Object getValue();

    void setValue(final Cell cell);
}
