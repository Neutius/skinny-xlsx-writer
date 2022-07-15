package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collection;

public class SkinnyWorkbookProvider implements XlsxWorkbookProvider {
    private final SXSSFWorkbook workbook;

    public SkinnyWorkbookProvider() {
        workbook = new SXSSFWorkbook();
    }

    public SkinnyWorkbookProvider(Collection<SheetContentSupplier> sheetContentSuppliers) {
        workbook = new SXSSFWorkbook();
        sheetContentSuppliers.forEach(this::addSheetToWorkbook);
    }

    public void addSheet(SheetContentSupplier sheetContentSupplier) {
        addSheetToWorkbook(sheetContentSupplier);
    }

    private void addSheetToWorkbook(SheetContentSupplier sheetContentSupplier) {
        SXSSFSheet sheet = workbook.createSheet();
        sheetContentSupplier.get().forEach(row -> addRowToSheet(row, sheet));
    }

    private void addRowToSheet(RowContentSupplier rowContent, SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        rowContent.get().forEach(cell -> addCellToRow(cell, row));
    }

    private void addCellToRow(String cellContent, SXSSFRow row) {
        SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
        cell.setCellValue(cellContent);
    }

    @Override
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

}
