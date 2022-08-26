package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.test.TestSheet;
import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.SoftAssertions;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWorkbookProviderTest {
	private static final String VALUE_1 = "value-1";
	private static final String SHEET_NAME = "sheet name";

	private static final String EMPTY_STRING = "";

	private static final String NULL_VALUE = null;
	private static final String SPACES = "    ";
	private static final String NEW_LINES = String.format("%n%n%n%n");
	private static final String TABS = "\t\t\t\t";

	private SkinnyWorkbookProvider testSubject;

	@Test
	void createInstance_isEmpty() {
		testSubject = new SkinnyWorkbookProvider();

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook).isNotNull().isEmpty();
	}

	@Test
	void createInstance_isXlsxWorkbook() {
		testSubject = new SkinnyWorkbookProvider();

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook).isNotInstanceOf(HSSFWorkbook.class);
	}

	@Test
	void addSheetToEmptyWorkbook_workbookHasSheet() {
		testSubject = new SkinnyWorkbookProvider();
		SkinnySheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier();
		sheetContentSupplier.addContentRow("value1");

		testSubject.addSheet(new SkinnySheetProvider(sheetContentSupplier));
		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook).isNotEmpty();
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("value1");
	}

	@Test
	void addSheetWithSeveralRowsAndColumns_allCellValuesAreInTheRightPlace() {
		testSubject = new SkinnyWorkbookProvider();
		SkinnySheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier();
		sheetContentSupplier.addContentRow("0-0-0", "0-0-1", "0-0-2", "0-0-3");
		sheetContentSupplier.addContentRow("0-1-0", "0-1-1", "0-1-2", "0-1-3");
		sheetContentSupplier.addContentRow("0-2-0", "0-2-1", "0-2-2", "0-2-3");
		sheetContentSupplier.addContentRow("0-3-0", "0-3-1", "0-3-2", "0-3-3");

		testSubject.addSheet(new SkinnySheetProvider(sheetContentSupplier));
		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("0-0-0");
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue()).isEqualTo("0-0-1");
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue()).isEqualTo("0-0-2");
		assertThat(workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue()).isEqualTo("0-0-3");

		assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo("0-1-0");
		assertThat(workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue()).isEqualTo("0-1-1");
		assertThat(workbook.getSheetAt(0).getRow(1).getCell(2).getStringCellValue()).isEqualTo("0-1-2");
		assertThat(workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue()).isEqualTo("0-1-3");

		assertThat(workbook.getSheetAt(0).getRow(2).getCell(0).getStringCellValue()).isEqualTo("0-2-0");
		assertThat(workbook.getSheetAt(0).getRow(2).getCell(1).getStringCellValue()).isEqualTo("0-2-1");
		assertThat(workbook.getSheetAt(0).getRow(2).getCell(2).getStringCellValue()).isEqualTo("0-2-2");
		assertThat(workbook.getSheetAt(0).getRow(2).getCell(3).getStringCellValue()).isEqualTo("0-2-3");

		assertThat(workbook.getSheetAt(0).getRow(3).getCell(0).getStringCellValue()).isEqualTo("0-3-0");
		assertThat(workbook.getSheetAt(0).getRow(3).getCell(1).getStringCellValue()).isEqualTo("0-3-1");
		assertThat(workbook.getSheetAt(0).getRow(3).getCell(2).getStringCellValue()).isEqualTo("0-3-2");
		assertThat(workbook.getSheetAt(0).getRow(3).getCell(3).getStringCellValue()).isEqualTo("0-3-3");
	}

	@Test
	void addSheetWithNullAndEmptyContentRows_emptyRowsAreAdded() {
		SheetContentSupplier contentSupplier = () -> Arrays.asList(
				() -> List.of(VALUE_1),
				() -> Collections.emptyList(),
				() -> null,
				null,
				() -> List.of(VALUE_1));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		SoftAssertions softly = new SoftAssertions();
		softly.assertThat(workbook.getSheetAt(0)).hasSize(5);
		softly.assertThat(workbook.getSheetAt(0).getRow(0)).hasSize(1);
		softly.assertThat(workbook.getSheetAt(0).getRow(1)).isNotNull().isEmpty();
		softly.assertThat(workbook.getSheetAt(0).getRow(2)).isNotNull().isEmpty();
		softly.assertThat(workbook.getSheetAt(0).getRow(3)).isNotNull().isEmpty();
		softly.assertThat(workbook.getSheetAt(0).getRow(4)).hasSize(1);

		softly.assertAll();
	}

	@Test
	void contentRowContainNullAndBlankValues_replacedWithEmptyStrings() {
		SheetContentSupplier contentSupplier = () -> List.of(() ->
				Arrays.asList(VALUE_1, NULL_VALUE, EMPTY_STRING, SPACES, TABS, NEW_LINES, VALUE_1));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		SoftAssertions softly = new SoftAssertions();
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(4).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(5).getStringCellValue()).isEqualTo(EMPTY_STRING);
		softly.assertThat(workbook.getSheetAt(0).getRow(0).getCell(6).getStringCellValue()).isEqualTo(VALUE_1);
		softly.assertAll();
	}

	@Test
	void useLambdasForSingleValue_valueIsPresent() {
		SkinnySheetProvider provider = new SkinnySheetProvider(() -> List.of(() -> List.of(VALUE_1)));
		testSubject = new SkinnyWorkbookProvider(List.of(provider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}

	@Test
	void addSeveralSheetsWithNoName_sheetsHaveUniqueNames() {
		SheetContentSupplier sameSheet = () -> List.of(() -> List.of(VALUE_1));
		SkinnySheetProvider provider = new SkinnySheetProvider(sameSheet);
		testSubject = new SkinnyWorkbookProvider(List.of(provider, provider, provider));

		Workbook workbook = testSubject.getWorkbook();
		String sheetName0 = workbook.getSheetAt(0).getSheetName();
		String sheetName1 = workbook.getSheetAt(1).getSheetName();
		String sheetName2 = workbook.getSheetAt(2).getSheetName();

		assertThat(sheetName0).isNotEqualTo(sheetName1);
		assertThat(sheetName0).isNotEqualTo(sheetName2);
		assertThat(sheetName1).isNotEqualTo(sheetName2);
	}

	@Test
	void addSheetWithName_sheetHasName() {
		SheetContentSupplier contentSupplier = () -> List.of(() -> List.of(VALUE_1));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
	}

	@Test
	void addSeveralSheetsWithTheSameName_sheetsHaveUniqueNames() {
		SheetContentSupplier sameSheet = () -> List.of(() -> List.of(VALUE_1));
		SkinnySheetProvider provider = new SkinnySheetProvider(sameSheet, SHEET_NAME);
		testSubject = new SkinnyWorkbookProvider(List.of(provider, provider, provider));

		Workbook workbook = testSubject.getWorkbook();
		String sheetName0 = workbook.getSheetAt(0).getSheetName();
		String sheetName1 = workbook.getSheetAt(1).getSheetName();
		String sheetName2 = workbook.getSheetAt(2).getSheetName();

		assertThat(sheetName0).isNotEqualTo(sheetName1);
		assertThat(sheetName0).isNotEqualTo(sheetName2);
		assertThat(sheetName1).isNotEqualTo(sheetName2);
	}

	@Test
	void cellValuesHaveDifferentWidth_sheetColumnsHaveDifferentWidth() {
		SheetContentSupplier sheet = () -> List.of(() -> List.of(
				"short",
				"medium sized",
				"relatively large piece of text"));
		testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(sheet, SHEET_NAME)));

		Workbook workbook = testSubject.getWorkbook();
		Sheet actualSheet = workbook.getSheetAt(0);

		assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
		assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
	}

	@Test
	void cellValuesHaveDifferentWidthOnlyAfter100Rows_sheetColumnsHaveSameWidth() {
		List<ContentRowSupplier> contentRowSupplierList = getListWithEmptyRows(100);
		contentRowSupplierList.add(() -> List.of(
				"short",
				"medium sized",
				"relatively large piece of text"));
		testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(() -> contentRowSupplierList, SHEET_NAME)));

		Workbook workbook = testSubject.getWorkbook();
		Sheet actualSheet = workbook.getSheetAt(0);

		assertThat(actualSheet.getColumnWidth(0)).isEqualTo(actualSheet.getColumnWidth(1));
		assertThat(actualSheet.getColumnWidth(1)).isEqualTo(actualSheet.getColumnWidth(2));
	}

	@Test
	void cellValuesHaveDifferentWidthOnlyAfter99Rows_sheetColumnsHaveDifferentWidth() {
		List<ContentRowSupplier> contentRowSupplierList = getListWithEmptyRows(99);
		contentRowSupplierList.add(() -> List.of(
				"short",
				"medium sized",
				"relatively large piece of text"));
		testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(() -> contentRowSupplierList, SHEET_NAME)));

		Workbook workbook = testSubject.getWorkbook();
		Sheet actualSheet = workbook.getSheetAt(0);

		assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
		assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
	}

	@Test
	void addSeveralSheets_columnsAreAutoSizedCorrectly() {
		List<String> differentLength = List.of(
				"short",
				"medium sized",
				"relatively large piece of text");
		List<ContentRowSupplier> sheetContent1 = List.of(() -> differentLength);
		List<ContentRowSupplier> sheetContent2 = getListWithEmptyRows(100);
		sheetContent2.add(() -> differentLength);
		List<ContentRowSupplier> sheetContent3 = getListWithEmptyRows(99);
		sheetContent3.add(() -> differentLength);
		List<ContentRowSupplier> sheetContent4 = getListWithEmptyRows(1);
		List<ContentRowSupplier> sheetContent5 = List.of(() -> List.of("Same", "Same", "Same"));

		testSubject = new SkinnyWorkbookProvider(List.of(new TestSheet(() -> sheetContent1), new TestSheet(() -> sheetContent2),
				new TestSheet(() -> sheetContent3), new TestSheet(() -> sheetContent4), new TestSheet(() -> sheetContent5)));

		Workbook workbook = testSubject.getWorkbook();
		Sheet actualSheet1 = workbook.getSheetAt(0);
		Sheet actualSheet2 = workbook.getSheetAt(1);
		Sheet actualSheet3 = workbook.getSheetAt(2);
		Sheet actualSheet4 = workbook.getSheetAt(3);
		Sheet actualSheet5 = workbook.getSheetAt(4);

		SoftAssertions softly = new SoftAssertions();
		softly.assertThat(actualSheet1.getColumnWidth(0)).isLessThan(actualSheet1.getColumnWidth(1));
		softly.assertThat(actualSheet1.getColumnWidth(1)).isLessThan(actualSheet1.getColumnWidth(2));
		softly.assertThat(actualSheet2.getColumnWidth(0)).isEqualTo(actualSheet2.getColumnWidth(1));
		softly.assertThat(actualSheet2.getColumnWidth(1)).isEqualTo(actualSheet2.getColumnWidth(2));
		softly.assertThat(actualSheet3.getColumnWidth(0)).isLessThan(actualSheet3.getColumnWidth(1));
		softly.assertThat(actualSheet3.getColumnWidth(1)).isLessThan(actualSheet3.getColumnWidth(2));
		softly.assertThat(actualSheet4.getColumnWidth(0)).isEqualTo(actualSheet4.getColumnWidth(1));
		softly.assertThat(actualSheet4.getColumnWidth(1)).isEqualTo(actualSheet4.getColumnWidth(2));
		softly.assertThat(actualSheet5.getColumnWidth(0)).isEqualTo(actualSheet5.getColumnWidth(1));
		softly.assertThat(actualSheet5.getColumnWidth(1)).isEqualTo(actualSheet5.getColumnWidth(2));
		softly.assertAll();
	}

	private static List<ContentRowSupplier> getListWithEmptyRows(int amountOfEmptyRows) {
		List<ContentRowSupplier> result = new ArrayList<>();
		for (int i = 1; i <= amountOfEmptyRows; i++) {
			result.add(Collections::emptyList);
		}
		return result;
	}

	@Test
	void addSheetsWithNullValues_sheetsHaveNamesAndNoContent() {
		SheetProvider sheet1 = new TestSheet(null, null, null);
		SheetProvider sheet2 = new TestSheet(null, () -> null, () -> null);
		testSubject = new SkinnyWorkbookProvider(List.of(sheet1, sheet2));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook).hasSize(2);
		assertThat(workbook.getSheetAt(0).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(0)).hasSize(0);
		assertThat(workbook.getSheetAt(1).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(1)).hasSize(0);
	}

	@Test
	void sheetsAreNull_sheetsHaveNamesAndNoContent() {
		List<SheetProvider> listContainingNull = new ArrayList<>();
		listContainingNull.add(null);
		testSubject = new SkinnyWorkbookProvider(listContainingNull);
		testSubject.addSheet(null);

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook).hasSize(2);
		assertThat(workbook.getSheetAt(0).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(0)).hasSize(0);
		assertThat(workbook.getSheetAt(1).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(1)).hasSize(0);
	}

	@Test
	void addSheetsWithEmptyAndBlankNames_sheetsHaveNames() {
		SheetProvider sheet1 = new TestSheet(EMPTY_STRING);
		SheetProvider sheet2 = new TestSheet(SPACES);
		SheetProvider sheet3 = new TestSheet(TABS);
		SheetProvider sheet4 = new TestSheet(NEW_LINES);
		testSubject = new SkinnyWorkbookProvider(List.of(sheet1, sheet2, sheet3, sheet4));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(1).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(2).getSheetName()).isNotNull().isNotBlank();
		assertThat(workbook.getSheetAt(3).getSheetName()).isNotNull().isNotBlank();
	}

	@Test
	void addSheetWithLongName_nameIsSnipped() {
		String longName = "abcdefghijklmnopqrstuvwxyz0123456789";
		SheetProvider sheet = new TestSheet(longName);
		testSubject = new SkinnyWorkbookProvider(List.of(sheet));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getSheetName()).isNotNull().isNotBlank().isEqualTo("abcdefghijklmnopqrstuvwxyz01234");
	}

	@Test
	void addSheetsWithLongNames_namesAreSnipped() {
		String longName = "Sheetnames cannot be longer than 31 characters";
		SheetProvider sheet1 = new TestSheet(longName);
		SheetProvider sheet2 = new TestSheet(longName);
		testSubject = new SkinnyWorkbookProvider(List.of(sheet1, sheet2));

		Workbook workbook = testSubject.getWorkbook();

		String longNameSnipped = longName.substring(0, 31);
		assertThat(workbook.getSheetAt(0).getSheetName()).isNotNull().isNotBlank().isEqualTo(longNameSnipped);
		assertThat(workbook.getSheetAt(1).getSheetName()).isNotNull().isNotBlank().isNotEqualTo(longNameSnipped);
	}

}
