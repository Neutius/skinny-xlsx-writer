package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxRowContentProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxSheetContentProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class SkinnyWorkbookProvider implements XlsxWorkbookProvider {
    private final SXSSFWorkbook workbook;

    public SkinnyWorkbookProvider() {
        workbook = new SXSSFWorkbook();
    }

    public void addSheet(XlsxSheetContentProvider sheetContentProvider) {
        SXSSFSheet sheet = workbook.createSheet();
        sheetContentProvider.getRowContentProviders().forEach(row -> addRowToSheet(row, sheet));
    }

    private void addRowToSheet(XlsxRowContentProvider rowContent, SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        rowContent.getRowContent().forEach(cell -> addCellToRow(cell, row));
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