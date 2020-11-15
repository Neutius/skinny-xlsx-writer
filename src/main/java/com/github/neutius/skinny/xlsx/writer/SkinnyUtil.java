package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.SimpleDateFormat;
import java.util.Date;

class SkinnyUtil {

    static final String EXTENSION = ".xlsx";

    private SkinnyUtil() {
        // nope
    }


    static String sanitizeFileName(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return "output-at-" + new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss").format(new Date());
        }
        return fileName;
    }

    static String sanitizeSheetName(String sheetName, Workbook workbook) {
        if (sheetName == null || sheetName.isBlank()) {
            return "Sheet_" + (workbook.getNumberOfSheets() + 1);
        }
        if (workbook.getSheet(sheetName) != null) {
            return sheetName + '_' + (workbook.getNumberOfSheets() + 1);
        }

        return sheetName;
    }

    static void adjustColumnSizesInCurrentSheet(Sheet currentSheet, int currentColumnAmount) {
        if (currentSheet == null) {
            return;
        }

        for (int index = 0; index < currentColumnAmount; index++) {
            currentSheet.autoSizeColumn(index);
        }
    }
}
