package com.github.neutius.skinny.xlsx.integration;

import com.github.neutius.skinny.xlsx.test.TestColumnHeaders;
import com.github.neutius.skinny.xlsx.test.TestSheet;
import com.github.neutius.skinny.xlsx.test.TestSheetContent;
import com.github.neutius.skinny.xlsx.writer.SkinnyColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.SkinnyContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.SkinnyFileWriter;
import com.github.neutius.skinny.xlsx.writer.SkinnySheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.SkinnySheetProvider;
import com.github.neutius.skinny.xlsx.writer.SkinnyWorkbookProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.assertj.core.api.SoftAssertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterIntegrationTest {
	private static final String LAMBDA_SHEET_NAME = "lambda-sheet";
	private static final String DEFAULT_IMPLEMENTATION_SHEET_NAME = "default-implementation";
	private static final String CUSTOM_IMPLEMENTATION_SHEET_NAME = "custom-implementation";
	private static final String NULL_VALUE_SHEET_NAME = null;
	private static final String CORNER_CASE_SHEET_NAME = "This sheet name is too long and will be snipped by Apache POI";

	private static final String VALUE_WITH_TABS = "we \t like \t tabs \t\t\t too \t much";
	private static final String VALUE_WITH_NEW_LINES = "we \n use \n new \n lines \n\n\n\n within a \n single \n value";
	private static final String VALUE_WITH_SPECIAL_CHARACTERS = " \\ \\ \" \" ; || ;; | , . \" ";
	private static final String HEART = "♥";
	private static final String ARROWS = "↨↑↓→←↔";

	private static final SheetProvider lambdaSheet = getLambdaSheet();
	private static final SheetProvider defaultImplementationSheet = getDefaultImplementationSheet();
	private static final SheetProvider customImplementationSheet = getCustomImplementationSheet();
	private static final SheetProvider nullValueSheet = getNullValueSheet();
	private static final SheetProvider cornerCaseSheet = getCornerCaseSheet();

	private static final List<SheetProvider> sheetProviderList = new ArrayList<>(List.of(
			lambdaSheet, defaultImplementationSheet, customImplementationSheet, nullValueSheet, cornerCaseSheet));

	static {
		sheetProviderList.add(null);
	}

	@Test
	void writeAndReadWorkbookToFileSystem(@TempDir File targetFolder) throws IOException, InvalidFormatException {
		SkinnyWorkbookProvider workbookProvider = new SkinnyWorkbookProvider(sheetProviderList);
		File outputFile = new File(targetFolder, "integration-test.xlsx");
		SkinnyFileWriter fileWriter = new SkinnyFileWriter();
		fileWriter.write(workbookProvider.getWorkbook(), outputFile);

		XSSFWorkbook actualWorkbook = new XSSFWorkbook(outputFile);

		assertWorkbookContent(actualWorkbook);
	}

	private void assertWorkbookContent(XSSFWorkbook actualWorkbook) {
		assertThat(actualWorkbook).isNotNull().isNotEmpty();

		SoftAssertions softly = new SoftAssertions();
		lambdaSheetHasAllContent(actualWorkbook, softly);
		defaultImplementationSheetHasAllContent(actualWorkbook, softly);
		customImplementationSheetHasAllContent(actualWorkbook, softly);
		nullValueSheetsHaveANameAndNoContent(actualWorkbook, softly);
		cornerCaseSheetHasAllContent(actualWorkbook, softly);
		softly.assertAll();
	}

	private static SheetProvider getLambdaSheet() {
		ColumnHeaderSupplier columnHeaderSupplier =
				() -> List.of("Header-0", "Header-1", "Header-2", "Header-3", "Header-4");
		SheetContentSupplier sheetContentSupplier = () -> List.of(
				() -> List.of("Value-0-0-0", "Value-0-0-1", "Value-0-0-2", "Value-0-0-3", "Value-0-0-4"),
				() -> List.of("Value-0-1-0", "Value-0-1-1", "Value-0-1-2", "Value-0-1-3", "Value-0-1-4"),
				() -> List.of("Value-0-2-0", "Value-0-2-1", "Value-0-2-2", "Value-0-2-3", "Value-0-2-4", "Value-0-2-5"),
				() -> List.of("Value-0-3-0", "Value-0-3-1", "Value-0-3-2", "Value-0-3-3"));

		return new SkinnySheetProvider(sheetContentSupplier, LAMBDA_SHEET_NAME, columnHeaderSupplier);
	}

	private static void lambdaSheetHasAllContent(XSSFWorkbook actualWorkbook, SoftAssertions softly) {
		XSSFSheet actualLambdaSheet = actualWorkbook.getSheet(LAMBDA_SHEET_NAME);

		softly.assertThat(actualLambdaSheet).isNotNull().isNotEmpty().hasSize(5);
		softly.assertThat(actualLambdaSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(3).getPhysicalNumberOfCells()).isEqualTo(6);
		softly.assertThat(actualLambdaSheet.getRow(4).getPhysicalNumberOfCells()).isEqualTo(4);
	}

	private static SheetProvider getDefaultImplementationSheet() {
		SkinnySheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier(
				new SkinnyContentRowSupplier("First content cell in first content row", "Value"),
				new SkinnyContentRowSupplier(List.of("Text", "Data")),
				new SkinnyContentRowSupplier()
		);
		sheetContentSupplier.addContentRow(List.of("Fourth row", "Text"));
		sheetContentSupplier.addContentRow("Fifth row", "Second content cell");
		sheetContentSupplier.addContentRow();
		sheetContentSupplier.addContentRow((String) null);
		sheetContentSupplier.addContentRow((Collection<String>) null);
		sheetContentSupplier.addContentRowSupplier(null);

		SkinnyContentRowSupplier duplicateContentRowSupplier = new SkinnyContentRowSupplier("Duplicate", "Row");
		sheetContentSupplier.addContentRowSupplier(duplicateContentRowSupplier);
		sheetContentSupplier.addContentRow(duplicateContentRowSupplier.get());

		SkinnyContentRowSupplier finalContentRowSupplier = new SkinnyContentRowSupplier("Last row");
		finalContentRowSupplier.addCellContent("final content cell");
		sheetContentSupplier.addContentRowSupplier(finalContentRowSupplier);

		SkinnyColumnHeaderSupplier headerSupplier = new SkinnyColumnHeaderSupplier("First column");
		headerSupplier.addColumnHeader("   ");
		headerSupplier.addColumnHeader(null);
		headerSupplier.addColumnHeader("Last column");

		return new SkinnySheetProvider(sheetContentSupplier, DEFAULT_IMPLEMENTATION_SHEET_NAME, headerSupplier);
	}

	private static void defaultImplementationSheetHasAllContent(XSSFWorkbook actualWorkbook, SoftAssertions softly) {
		XSSFSheet actualSheet = actualWorkbook.getSheet(DEFAULT_IMPLEMENTATION_SHEET_NAME);

		softly.assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(13);
		softly.assertThat(actualSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(4);
		softly.assertThat(actualSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("First column");
		softly.assertThat(actualSheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(0).getCell(3).getStringCellValue()).isEqualTo("Last column");

		softly.assertThat(actualSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(3).getPhysicalNumberOfCells()).isEqualTo(0);
		softly.assertThat(actualSheet.getRow(4).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(5).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(6).getPhysicalNumberOfCells()).isEqualTo(0);
		softly.assertThat(actualSheet.getRow(7).getPhysicalNumberOfCells()).isEqualTo(1);
		softly.assertThat(actualSheet.getRow(8).getPhysicalNumberOfCells()).isEqualTo(0);
		softly.assertThat(actualSheet.getRow(9).getPhysicalNumberOfCells()).isEqualTo(0);
		softly.assertThat(actualSheet.getRow(10).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(11).getPhysicalNumberOfCells()).isEqualTo(2);
		softly.assertThat(actualSheet.getRow(12).getPhysicalNumberOfCells()).isEqualTo(2);
	}

	private static SheetProvider getCustomImplementationSheet() {
		ColumnHeaderSupplier columnHeaderSupplier = new TestColumnHeaders();
		SheetContentSupplier sheetContentSupplier = new TestSheetContent();

		return new TestSheet(CUSTOM_IMPLEMENTATION_SHEET_NAME, columnHeaderSupplier, sheetContentSupplier);
	}

	private static void customImplementationSheetHasAllContent(XSSFWorkbook actualWorkbook, SoftAssertions softly) {
		XSSFSheet actualSheet = actualWorkbook.getSheet(CUSTOM_IMPLEMENTATION_SHEET_NAME);

		softly.assertThat(actualSheet).hasSize(2);
		softly.assertThat(actualSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(4);
		softly.assertThat(actualSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("First column");
		softly.assertThat(actualSheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(0).getCell(3).getStringCellValue()).isEqualTo("Last column");
		softly.assertThat(actualSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(4);
		softly.assertThat(actualSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("First content cell");
		softly.assertThat(actualSheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("");
		softly.assertThat(actualSheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("Last content cell");
	}

	private static SheetProvider getNullValueSheet() {
		return new TestSheet(NULL_VALUE_SHEET_NAME, null, null);
	}

	private static void nullValueSheetsHaveANameAndNoContent(XSSFWorkbook actualWorkbook, SoftAssertions softly) {
		List<XSSFSheet> nullValueSheets = findNullValueSheets(actualWorkbook);

		softly.assertThat(nullValueSheets).hasSize(2);

		for (XSSFSheet nullValueSheet : nullValueSheets) {
			softly.assertThat(nullValueSheet).isNotNull().isEmpty();
			softly.assertThat(nullValueSheet.getSheetName()).isNotNull().isNotBlank();
			softly.assertThatThrownBy(() -> nullValueSheet.getRow(0).getCell(0))
					.isInstanceOf(NullPointerException.class);
		}
	}

	private static List<XSSFSheet> findNullValueSheets(XSSFWorkbook actualWorkbook) {
		return StreamSupport.stream(actualWorkbook.spliterator(), false)
				.filter(sheet -> !(LAMBDA_SHEET_NAME.equals(sheet.getSheetName())))
				.filter(sheet -> !(DEFAULT_IMPLEMENTATION_SHEET_NAME.equals(sheet.getSheetName())))
				.filter(sheet -> !(CUSTOM_IMPLEMENTATION_SHEET_NAME.equals(sheet.getSheetName())))
				.filter(sheet -> !(CORNER_CASE_SHEET_NAME.startsWith(sheet.getSheetName())))
				.filter(sheet -> sheet instanceof XSSFSheet)
				.map(sheet -> (XSSFSheet) sheet)
				.collect(Collectors.toList());
	}

	private static SheetProvider getCornerCaseSheet() {
		ColumnHeaderSupplier columnHeaderSupplier = new SkinnyColumnHeaderSupplier(
				List.of("entry0", "1", "?", "Mariënberg", "Curaçao"));

		List<String> firstRow = List.of(VALUE_WITH_TABS, VALUE_WITH_NEW_LINES, VALUE_WITH_SPECIAL_CHARACTERS, HEART, ARROWS);
		List<String> secondRow = List.of("entry1", "false", "true", "null");
		List<String> thirdRow = List.of("entry2", "", "", "", "", "sixth      column");

		SheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier(
				List.of(() -> firstRow, () -> secondRow, () -> thirdRow));

		return new TestSheet(CORNER_CASE_SHEET_NAME, columnHeaderSupplier, sheetContentSupplier);
	}

	private static void cornerCaseSheetHasAllContent(XSSFWorkbook actualWorkbook, SoftAssertions softly) {
		XSSFSheet absentSheet = actualWorkbook.getSheet(CORNER_CASE_SHEET_NAME);
		XSSFSheet actualSheet = actualWorkbook.getSheet(CORNER_CASE_SHEET_NAME.substring(0, 31));

		softly.assertThat(absentSheet).isNull();
		softly.assertThat(actualSheet).hasSize(4);
		assertCellValue(actualSheet, softly, 0, 0, "entry0");
		assertCellValue(actualSheet, softly, 0, 1, "1");
		assertCellValue(actualSheet, softly, 0, 2, "?");
		assertCellValue(actualSheet, softly, 0, 3, "Mariënberg");
		assertCellValue(actualSheet, softly, 0, 4, "Curaçao");
		assertCellValue(actualSheet, softly, 1, 0, VALUE_WITH_TABS);
		assertCellValue(actualSheet, softly, 1, 1, VALUE_WITH_NEW_LINES);
		assertCellValue(actualSheet, softly, 1, 2, VALUE_WITH_SPECIAL_CHARACTERS);
		assertCellValue(actualSheet, softly, 1, 3, HEART);
		assertCellValue(actualSheet, softly, 1, 4, ARROWS);

		assertCellValue(actualSheet, softly, 2, 0, "entry1");
		assertCellValue(actualSheet, softly, 2, 1, "false");
		assertCellValue(actualSheet, softly, 2, 2, "true");
		assertCellValue(actualSheet, softly, 2, 3, "null");

		assertCellValue(actualSheet, softly, 3, 0, "entry2");
		assertCellValue(actualSheet, softly, 3, 1, "");
		assertCellValue(actualSheet, softly, 3, 2, "");
		assertCellValue(actualSheet, softly, 3, 3, "");
		assertCellValue(actualSheet, softly, 3, 4, "");
		assertCellValue(actualSheet, softly, 3, 5, "sixth      column");
	}

	private static void assertCellValue(XSSFSheet actualSheet, SoftAssertions softly, int row, int cell, String expected) {
		softly.assertThat(actualSheet.getRow(row).getCell(cell).getStringCellValue()).isEqualTo(expected);
	}

	// TODO for coverage: call SkinnySheetContentSupplier()
	// TODO for coverage: call SkinnyColumnHeaderSupplier()

}
