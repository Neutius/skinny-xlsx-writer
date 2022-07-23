package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collection;

public class SkinnyWorkbookProvider implements XlsxWorkbookProvider {
    private final SXSSFWorkbook workbook;

    @Override
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public SkinnyWorkbookProvider() {
        workbook = new SXSSFWorkbook();
    }

    public SkinnyWorkbookProvider(Collection<SheetProvider> sheetProviders) {
        workbook = new SXSSFWorkbook();
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
        sheetProvider.getSheetContentSupplier().get().forEach(row -> addRowToSheet(row, sheet));
    }

    private void addSheetToWorkbook(SheetContentSupplier sheetContentSupplier) {
        SXSSFSheet sheet = createSheet("");
        sheetContentSupplier.get().forEach(row -> addRowToSheet(row, sheet));
    }

    private SXSSFSheet createSheet(String sheetName) {
        boolean hasName = sheetName != null && !(sheetName.isBlank());
        return hasName ? workbook.createSheet(sanitizeSheetName(sheetName)) : workbook.createSheet();
    }

    private String sanitizeSheetName(String sheetName) {
        return isUnique(sheetName) ? sheetName : sheetName + "-" + workbook.getNumberOfSheets();
    }

    private boolean isUnique(String sheetName) {
        for (Sheet sheet : workbook) {
            if (sheet.getSheetName().equals(sheetName)) {
                return false;
            }
        }

        return true;
    }

    private void addRowToSheet(RowContentSupplier rowContent, SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        rowContent.get().forEach(cell -> addCellToRow(cell, row));
    }

    private void addCellToRow(String cellContent, SXSSFRow row) {
        SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
        cell.setCellValue(cellContent);
    }

}
