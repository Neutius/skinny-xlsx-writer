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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.assertj.core.api.SoftAssertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.Collection;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterIntegrationTest {
	private static final String LAMBDA_SHEET_NAME = "lambda-sheet";
	private static final String DEFAULT_IMPLEMENTATION_SHEET_NAME = "default-implementation";
	private static final String CUSTOM_IMPLEMENTATION_SHEET_NAME = "custom-implementation";
	private static final String NULL_VALUE_SHEET_NAME = null;
	private static final String CORNER_CASE_SHEET_NAME = "This sheet name is too long and will be snipped by Apache POI";

	private static XSSFWorkbook actualWorkbook;

	@BeforeAll
	static void writeAndReadWorkbookToFileSystem(@TempDir File targetFolder) throws IOException, InvalidFormatException {
		SheetProvider lambdaSheet = getLambdaSheet();
		SheetProvider defaultImplementationSheet = getDefaultImplementationSheet();
		SheetProvider customImplementationSheet = getCustomImplementationSheet();
		SheetProvider nullValueSheet = getNullValueSheet();
		SheetProvider cornerCaseSheet = getCornerCaseSheet();

		SkinnyWorkbookProvider workbookProvider = new SkinnyWorkbookProvider(List.of(
				lambdaSheet, defaultImplementationSheet, customImplementationSheet, nullValueSheet, cornerCaseSheet));

		File outputFile = writeToFile(targetFolder, workbookProvider.getWorkbook());

		actualWorkbook = new XSSFWorkbook(outputFile);
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

	@Test
	void lambdaSheetHasAllContent() {
		XSSFSheet actualLambdaSheet = actualWorkbook.getSheet(LAMBDA_SHEET_NAME);

		SoftAssertions softly = new SoftAssertions();

		softly.assertThat(actualLambdaSheet).isNotNull().isNotEmpty().hasSize(5);
		softly.assertThat(actualLambdaSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(5);
		softly.assertThat(actualLambdaSheet.getRow(3).getPhysicalNumberOfCells()).isEqualTo(6);
		softly.assertThat(actualLambdaSheet.getRow(4).getPhysicalNumberOfCells()).isEqualTo(4);

		softly.assertAll();
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

	@Test
	void defaultImplementationSheetHasAllContent() {
		XSSFSheet actualSheet = actualWorkbook.getSheet(DEFAULT_IMPLEMENTATION_SHEET_NAME);

		SoftAssertions softly = new SoftAssertions();

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

		softly.assertAll();
	}

	private static SheetProvider getCustomImplementationSheet() {
		ColumnHeaderSupplier columnHeaderSupplier = new TestColumnHeaders();
		SheetContentSupplier sheetContentSupplier = new TestSheetContent();

		return new TestSheet(CUSTOM_IMPLEMENTATION_SHEET_NAME, columnHeaderSupplier, sheetContentSupplier);
	}

	@Test
	void customImplementationSheetHasAllContent() {
		XSSFSheet actualSheet = actualWorkbook.getSheet(CUSTOM_IMPLEMENTATION_SHEET_NAME);

		SoftAssertions softly = new SoftAssertions();

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

		softly.assertAll();
	}

	private static SheetProvider getNullValueSheet() {
		return new TestSheet(NULL_VALUE_SHEET_NAME, null, null);
	}

	private static SheetProvider getCornerCaseSheet() {
		return new TestSheet(CORNER_CASE_SHEET_NAME, null, null);
	}

	// TODO for coverage: call SkinnySheetContentSupplier()
	// TODO for coverage: call SkinnySheetContentSupplier(Collection<ContentRowSupplier> initialContent)
	// TODO for coverage: call SkinnyColumnHeaderSupplier()
	// TODO for coverage: call SkinnyColumnHeaderSupplier(Collection<String> initialContent)

	private static File writeToFile(File targetFolder, SXSSFWorkbook workbook) {
		File outputFile = new File(targetFolder, "integration-test.xlsx");
		SkinnyFileWriter fileWriter = new SkinnyFileWriter();
		fileWriter.write(workbook, outputFile);
		return outputFile;
	}

	@Test
	void fileIsReadable() {
		assertThat(actualWorkbook).isNotNull().isNotEmpty();
	}

}
