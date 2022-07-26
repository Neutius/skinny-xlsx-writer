package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

public class SkinnyWorkbookProvider implements XlsxWorkbookProvider {
    private final SXSSFWorkbook workbook = new SXSSFWorkbook();

    @Override
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public SkinnyWorkbookProvider() {
    }

    public SkinnyWorkbookProvider(Collection<SheetProvider> sheetProviders) {
        sheetProviders.forEach(this::addSheetToWorkbook);
    }

    /*
    The separate "addSheet" method consumes a SheetContentSupplier, but the constructor consumes SheetProvider instances.
    This is inconsistent. Perhaps add an addSheet overload that consumes a SheetProvider instance?
    Or change this method and remove support for SheetContentSupplier in this API?
    Perhaps split this class into a "simple" and "deluxe" version - without and with sheet names and column headers?
    GvdNL 23-07-2022
    */
    public void addSheet(SheetContentSupplier sheetContentSupplier) {
        addSheetToWorkbook(sheetContentSupplier);
    }

    private void addSheetToWorkbook(SheetProvider sheetProvider) {
        SXSSFSheet sheet = createSheet(sheetProvider.getSheetName());
        fillSheet(sheetProvider.getSheetContentSupplier(), sheet);
    }

    private void addSheetToWorkbook(SheetContentSupplier sheetContentSupplier) {
        SXSSFSheet sheet = createSheet("");
        fillSheet(sheetContentSupplier, sheet);
    }

    private SXSSFSheet createSheet(String sheetName) {
        boolean isValidName = sheetName != null && !(sheetName.isBlank());
        return isValidName ? workbook.createSheet(sanitizeSheetName(sheetName)) : workbook.createSheet();
    }

    private String sanitizeSheetName(String sheetName) {
        return isUnique(sheetName) ? sheetName : sheetName + "-" + workbook.getNumberOfSheets();
    }

    private boolean isUnique(String sheetName) {
        Set<String> sheetNamesInWorkbook = new HashSet<>();
        workbook.forEach(sheet -> sheetNamesInWorkbook.add(sheet.getSheetName()));
        return !sheetNamesInWorkbook.contains(sheetName);
    }

    private static void fillSheet(SheetContentSupplier sheetContentSupplier, SXSSFSheet sheet) {
        sheetContentSupplier.get().forEach(row -> addRowToSheet(row, sheet));
        if (sheet.getPhysicalNumberOfRows() < 100) {
            autoSizeColumns(sheet);
        }
    }

    private static void addRowToSheet(ContentRowSupplier contentRow, SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        contentRow.get().forEach(cell -> addCellToRow(cell, row));
        if (sheet.getPhysicalNumberOfRows() == 100) {
            autoSizeColumns(sheet);
        }
    }

    private static void addCellToRow(String cellContent, SXSSFRow row) {
        SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
        cell.setCellValue(cellContent);
    }

    private static void autoSizeColumns(SXSSFSheet sheet) {
        sheet.trackAllColumnsForAutoSizing();
        autoSizeColumns(sheet, getCurrentAmountOfColumns(sheet));
        sheet.untrackAllColumnsForAutoSizing();
    }

    private static int getCurrentAmountOfColumns(Sheet sheet) {
        int columnAmount = 0;
        for (Row row : sheet) {
            columnAmount = Math.max(row.getPhysicalNumberOfCells(), columnAmount);
        }
        return columnAmount;
    }

    private static void autoSizeColumns(Sheet sheet, int columnAmount) {
        for (int index = 0; index < columnAmount; index++) {
            sheet.autoSizeColumn(index);
        }
    }

}
