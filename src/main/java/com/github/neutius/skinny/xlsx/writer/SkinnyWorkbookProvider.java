package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.UUID;

public class SkinnyWorkbookProvider implements XlsxWorkbookProvider {
	private final SXSSFWorkbook workbook = new SXSSFWorkbook();
	private final SheetNameHandler nameHandler;

	@Override
	public SXSSFWorkbook getWorkbook() {
		return workbook;
	}

	public SkinnyWorkbookProvider() {
		nameHandler = new SheetNameHandler(workbook);
	}

	public SkinnyWorkbookProvider(Collection<SheetProvider> sheetProviders) {
		nameHandler = new SheetNameHandler(workbook);
		sheetProviders.forEach(this::addSheetToWorkbook);
	}

	public void addSheet(SheetProvider sheetProvider) {
		addSheetToWorkbook(sheetProvider);
	}

	private void addSheetToWorkbook(SheetProvider sheetProvider) {
		SXSSFSheet sheet = createSheet(sheetProvider);
		addColumnHeaders(sheetProvider, sheet);
		fillSheet(sheetProvider, sheet);
	}

	private SXSSFSheet createSheet(SheetProvider sheetProvider) {
		return sheetProvider == null || sheetProvider.getSheetName() == null || sheetProvider.getSheetName().isBlank()
				? workbook.createSheet()
				: workbook.createSheet(nameHandler.sanitize(sheetProvider.getSheetName()));
	}

	private static void addColumnHeaders(SheetProvider sheetProvider, SXSSFSheet sheet) {
		if (columnHeadersAreProvided(sheetProvider)) {
			addRowToSheet(sheetProvider.getColumnHeaderSupplier().get(), sheet);
			applyColumnHeaderFormattingToFirstRow(sheet);
			sheet.createFreezePane(0, 1);
			sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, sheetProvider.getColumnHeaderSupplier().get().size() - 1));
		}
	}

	private static boolean columnHeadersAreProvided(SheetProvider sheetProvider) {
		return sheetProvider != null
				&& sheetProvider.getColumnHeaderSupplier() != null
				&& sheetProvider.getColumnHeaderSupplier().get() != null
				&& !(sheetProvider.getColumnHeaderSupplier().get().isEmpty());
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

	private static void fillSheet(SheetProvider sheetProvider, SXSSFSheet sheet) {
		if (sheetProvider == null || sheetProvider.getSheetContentSupplier() == null
				|| sheetProvider.getSheetContentSupplier().get() == null) {
			return;
		}
		sheetProvider.getSheetContentSupplier().get().forEach(row -> addRowToSheet(sanitizeRow(row).get(), sheet));
		if (sheet.getPhysicalNumberOfRows() < 100) {
			autoSizeColumns(sheet);
		}
	}

	private static ContentRowSupplier sanitizeRow(ContentRowSupplier row) {
		if (row == null || row.get() == null) {
			return Collections::emptyList;
		}
		return row;
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
		cell.setCellValue(sanitizeCellContent(cellContent));
	}

	private static String sanitizeCellContent(String cellContent) {
		if (cellContent == null || cellContent.isBlank()) {
			return "";
		}
		return cellContent;
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

	private static class SheetNameHandler {
		private static final int MAX_LENGTH = 31;
		private final Workbook workbook;
		private Set<String> sheetNamesInWorkbook;

		private SheetNameHandler(Workbook workbook) {
			this.workbook = workbook;
		}

		private String sanitize(String sheetName) {
			sheetNamesInWorkbook = new HashSet<>();
			workbook.forEach(sheet -> sheetNamesInWorkbook.add(sheet.getSheetName()));
			return isUnique(sheetName) ? sheetName : createUniqueSheetName(sheetName);
		}

		private boolean isUnique(String sheetName) {
			String snippedSheetName = sheetName.length() <= MAX_LENGTH ? sheetName : sheetName.substring(0, MAX_LENGTH);
			return !sheetNamesInWorkbook.contains(snippedSheetName);
		}

		private String createUniqueSheetName(String sheetName) {
			String numberedSheetName = sheetName + "-" + workbook.getNumberOfSheets();
			if (isUnique(numberedSheetName)) {
				return numberedSheetName;
			}
			return padWithRandomCharacters(sheetName);
		}

		private String padWithRandomCharacters(String sheetName) {
			String shortenedSheetName = sheetName.length() <= 22 ? sheetName : sheetName.substring(0, 22);
			String paddedSheetName = shortenedSheetName + "-" + UUID.randomUUID().toString().substring(0, 8);
			return isUnique(paddedSheetName) ? paddedSheetName : padWithRandomCharacters(sheetName);
		}

	}

}
