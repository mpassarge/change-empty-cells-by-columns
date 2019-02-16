package com.passargecorp.excel.workbook;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Work {

    private final static int DEFAULT_SHEET = 0;

    private static Workbook workbook;
    private static Sheet sheet;

    private Work(){/**/}

    public static void runner(final String columnCoordinate, final String filePath, final TypeOfOutput type) {

        final boolean bookLoaded = loadBook(filePath);
        if(bookLoaded){

            final int colNumber = CellReference.convertColStringToIndex(columnCoordinate);

            sheet = workbook.getSheetAt(DEFAULT_SHEET);
            if(columnInSheet(colNumber)) {
                for(Row row : sheet){
                    final Cell cell = row.getCell(colNumber);
                    if(cell.getCellType().equals(CellType.BLANK)){
                        cell.setCellValue("");
                    }
                }
            } else{
                // TODO: 2/15/19 Log error message
            }
        }
        try {
            FileUtils.moveFile(new File(filePath), new File(filePath+".old"));
            FileOutputStream fo = new FileOutputStream(new File(filePath+".new"));
            workbook.write(fo);
            fo.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean columnInSheet(final int colNumber){
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

}