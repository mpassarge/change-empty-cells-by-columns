package com.passargecorp.excel.workbook;

import com.passargecorp.excel.workbook.output.OutputInterface;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Work {

    private final static int DEFAULT_SHEET = 0;

    private static Workbook workbook;
    private static Sheet sheet;

    private Work() {/**/}

    private static boolean columnInSheet(final int colNumber) {
        return colNumber < sheet.getLastRowNum();
    }

    private static boolean loadBook(String filePath) {
        try {
            workbook = WorkbookFactory.create(new File(filePath));
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            return true;
        } catch (IOException e) {
            // TODO: 2/15/19 Log error message
            return false;
        }
    }

    public static void runner(final String columnCoordinate, final String filePath, final OutputInterface output) {

        final boolean bookLoaded = loadBook(filePath);
        if (bookLoaded) {

            final int colNumber = CellReference.convertColStringToIndex(columnCoordinate);

            sheet = workbook.getSheetAt(DEFAULT_SHEET);
            if (columnInSheet(colNumber)) {
                for (Row row : sheet) {
                    final Cell cell = row.getCell(colNumber);
                    if (cell.getCellType().equals(CellType.BLANK)) {
                        output.setValue(cell);
                    }
                }
            } else {
                // TODO: 2/15/19 Log error message
            }
        }
        outputWorkbook(filePath);
    }

    private static void outputWorkbook(final String filePath) {
        try {
            saveOldWorkbook(filePath);
            saveNewWorkbook(filePath);
        } catch (IOException e) {
            // TODO: 2/16/19 log error message
        }
    }

    private static void saveOldWorkbook(final String filePath) throws IOException {
        FileUtils.moveFile(new File(filePath), new File(filePath + ".old"));
    }

    private static void saveNewWorkbook(final String filePath) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath + ".new"));
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

}