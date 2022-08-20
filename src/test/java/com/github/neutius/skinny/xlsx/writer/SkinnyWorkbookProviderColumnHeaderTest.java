package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.assertj.core.api.SoftAssertions;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWorkbookProviderColumnHeaderTest {
	private static final String HEADER_1 = "Header-1";
	private static final String VALUE_1 = "value-1";
	private static final SheetContentSupplier CONTENT_SUPPLIER = () -> List.of(() -> List.of(VALUE_1));
	private static final String SHEET_NAME = "sheet name";

	private static final String EMPTY_STRING = "";

	private static final String NULL_VALUE = null;
	private static final String SPACES = "    ";
	private static final String NEW_LINES = String.format("%n%n%n%n");
	private static final String TABS = "\t\t\t\t";

	private SkinnyWorkbookProvider testSubject;

	@Test
	void addColumnHeader_isPresent() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(HEADER_1);
	}

	@Test
	void addColumnHeaderAndContentRow_bothArePresent() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(2);
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(HEADER_1);
		assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}

	@Test
	void columnHeaderSupplierReturnsEmptyList_noColumnHeadersAdded() {
		ColumnHeaderSupplier headerSupplier = () -> Collections.emptyList();
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(1);
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}

	@Test
	void columnHeaderSupplierReturnsNull_noColumnHeadersAdded() {
		ColumnHeaderSupplier headerSupplier = () -> null;
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(1);
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}

	@Test
	void columnHeaderSupplierIsNull_noColumnHeadersAdded() {
		ColumnHeaderSupplier headerSupplier = null;
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(1);
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}

	@Test
	void columnHeadersContainNullAndBlankValues_replacedWithEmptyStrings() {
		ColumnHeaderSupplier headerSupplier = () ->
				Arrays.asList(HEADER_1, NULL_VALUE, EMPTY_STRING, SPACES, TABS, NEW_LINES, HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		SoftAssertions softly = new SoftAssertions();
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(HEADER_1);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(4).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(5).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(6).getStringCellValue()).isEqualTo(HEADER_1);
		softly.assertAll();
	}

	@Test
	void addColumnHeader_fontIsBold() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();

		SXSSFCell cell = workbook.getSheetAt(0).getRow(0).getCell(0);
		XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
		XSSFFont font = cellStyle.getFont();
		assertThat(font).isNotNull();
		assertThat(font.getBold()).isTrue();
	}

	@Test
	void addColumnHeaderAndContentRow_contentRowFontIsNotBold() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();

		SXSSFCell cell = workbook.getSheetAt(0).getRow(1).getCell(0);
		XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
		XSSFFont font = cellStyle.getFont();
		assertThat(font).isNotNull();
		assertThat(font.getBold()).isFalse();
	}

	@Test
	void noColumnHeader_firstRowFontIsNotBold() {
		ColumnHeaderSupplier headerSupplier = () -> null;
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();

		SXSSFCell cell = workbook.getSheetAt(0).getRow(0).getCell(0);
		XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
		XSSFFont font = cellStyle.getFont();
		assertThat(font).isNotNull();
		assertThat(font.getBold()).isFalse();
	}

	@Test
	void addColumnHeader_freezePaneIsApplied() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();

		PaneInformation paneInformation = workbook.getSheetAt(0).getPaneInformation();
		assertThat(paneInformation).isNotNull();
		assertThat(paneInformation.isFreezePane()).isTrue();
		assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
		assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
	}

	@Test
	void noColumnHeader_freezePaneIsNotApplied() {
		ColumnHeaderSupplier headerSupplier = () -> null;
		SheetProvider sheetProvider = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();

		PaneInformation paneInformation = workbook.getSheetAt(0).getPaneInformation();
		assertThat(paneInformation).isNull();
	}

	@Test
	void columnHeadersHaveSameWidth_sheetColumnsHaveSameWidth() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(
				"Same width",
				"Same width",
				"Same width");
		SheetContentSupplier contentSupplier = () -> List.of(() -> List.of(
				"short",
				"short",
				"short"));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();
		Sheet actualSheet = workbook.getSheetAt(0);

		assertThat(actualSheet.getColumnWidth(0)).isEqualTo(actualSheet.getColumnWidth(1));
		assertThat(actualSheet.getColumnWidth(1)).isEqualTo(actualSheet.getColumnWidth(2));
	}

	@Test
	void columnHeadersHaveDifferentWidth_sheetColumnsHaveDifferentWidth() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(
				"short",
				"Same width",
				"Relatively long column header text");
		SheetContentSupplier contentSupplier = () -> List.of(() -> List.of(
				"short",
				"short",
				"short"));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		SXSSFWorkbook workbook = testSubject.getWorkbook();
		Sheet actualSheet = workbook.getSheetAt(0);

		assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
		assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
	}


	/*
	TODO add tests and functionality - GvdNL 26-07-2022
	- configuration options for bold font, freeze pane and filter?

	Maybe here or somewhere else?
	- cellValuesHaveDifferentWidthOnlyAfter99ContentRows_sheetColumnsHaveSameWidth
	- cellValuesHaveDifferentWidthOnlyAfter98ContentRows_sheetColumnsHaveDifferentWidth

	 */


}
