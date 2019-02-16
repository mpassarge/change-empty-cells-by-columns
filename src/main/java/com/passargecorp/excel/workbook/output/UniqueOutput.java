package com.passargecorp.excel.workbook.output;

public class UniqueOutput extends StringOutput implements OutputInterface {

    public UniqueOutput(final String addedUniqueValue) {
        super("Null " + addedUniqueValue);
    }
}
