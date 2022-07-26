package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.util.Collection;
import java.util.HashSet;
import java.util.List;
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

    public void addSheet(SheetProvider sheetProvider) {
        addSheetToWorkbook(sheetProvider);
    }

    private void addSheetToWorkbook(SheetProvider sheetProvider) {
        SXSSFSheet sheet = createSheet(sheetProvider.getSheetName());
        addColumnHeaders(sheetProvider, sheet);
        fillSheet(sheetProvider.getSheetContentSupplier(), sheet);
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

    private static void addColumnHeaders(SheetProvider sheetProvider, SXSSFSheet sheet) {
        if (columnHeadersAreProvided(sheetProvider.getColumnHeaderSupplier())) {
            addRowToSheet(sheetProvider.getColumnHeaderSupplier().get(), sheet);
            applyColumnHeaderFormattingToFirstRow(sheet);
        }
    }

    private static void applyColumnHeaderFormattingToFirstRow(SXSSFSheet sheet) {
        CellStyle columnHeaderCellStyle = createColumnHeaderCellStyle(sheet.getWorkbook());
        for (Cell cell : sheet.getRow(0)) {
            cell.setCellStyle(columnHeaderCellStyle);
        }
    }

    private static CellStyle createColumnHeaderCellStyle(SXSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(getBoldFont(workbook));
        style.setWrapText(false);
        return style;
    }

    private static Font getBoldFont(SXSSFWorkbook workbook) {
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        return boldFont;
    }

    private static boolean columnHeadersAreProvided(ColumnHeaderSupplier columnHeaderSupplier) {
        return columnHeaderSupplier != null && columnHeaderSupplier.get() != null && !(columnHeaderSupplier.get().isEmpty());
    }

    private static void fillSheet(SheetContentSupplier sheetContentSupplier, SXSSFSheet sheet) {
        sheetContentSupplier.get().forEach(row -> addRowToSheet(row.get(), sheet));
        if (sheet.getPhysicalNumberOfRows() < 100) {
            autoSizeColumns(sheet);
        }
    }

    private static void addRowToSheet(List<String> cellValues, SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
        cellValues.forEach(cell -> addCellToRow(cell, row));
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
        for (int index = 0; index < getCurrentAmountOfColumns(sheet); index++) {
            sheet.autoSizeColumn(index);
        }
        sheet.untrackAllColumnsForAutoSizing();
    }

    private static int getCurrentAmountOfColumns(Sheet sheet) {
        int columnAmount = 0;
        for (Row row : sheet) {
            columnAmount = Math.max(row.getPhysicalNumberOfCells(), columnAmount);
        }
        return columnAmount;
    }

}
