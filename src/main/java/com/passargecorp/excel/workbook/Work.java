package com.passargecorp.excel.workbook;

import com.passargecorp.excel.workbook.output.OutputInterface;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Work {

    private final static int DEFAULT_SHEET = 0;
    private static Logger LOG = LoggerFactory.getLogger(Work.class);

    private static Workbook workbook;
    private static Sheet sheet;

    private Work() {/**/}

    private static boolean columnInSheet(final int colNumber) {
        return colNumber < sheet.getRow(0).getLastCellNum();
    }

    private static boolean loadBook(String filePath) {
        LOG.info("Loading work book from file \"{}\"", filePath);
        try {
            workbook = WorkbookFactory.create(new File(filePath));
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            return true;
        } catch (IOException e) {
            LOG.error("Could not read from file \"{}\"", filePath);
            return false;
        }
    }

    public static void runner(final String columnCoordinate, final String filePath, final OutputInterface output) {

        final boolean bookLoaded = loadBook(filePath);
        if (bookLoaded) {
            LOG.info("Workbook loaded...");
            final int colNumber = CellReference.convertColStringToIndex(columnCoordinate);
            sheet = workbook.getSheetAt(DEFAULT_SHEET);
            if (columnInSheet(colNumber)) {
                LOG.info("Editing Column {} with {}", columnCoordinate, output.getValue());
                for (Row row : sheet) {
                    final Cell cell = row.getCell(colNumber);
                    if (cell.getCellType().equals(CellType.BLANK)) {
                        output.setValue(cell);
                    }
                }
            }
            outputWorkbook(filePath);
        }
    }

    private static void outputWorkbook(final String filePath) {
        LOG.info("Saving workbook");
        try {
            saveOldWorkbook(filePath);
            saveNewWorkbook(filePath);
        } catch (IOException e) {
            LOG.error("Error saving workbook at \"{}\"", filePath);
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