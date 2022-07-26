package com.github.neutius.skinny.xlsx.writer;

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
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterIntegrationTest {
	private static final String LAMBDA_SHEET_NAME = "lambda-sheet";

	private static XSSFWorkbook actualWorkbook;

	@BeforeAll
	static void writeAndReadWorkbookToFileSystem(@TempDir File targetFolder) throws IOException, InvalidFormatException {
		SheetProvider lambdaSheetProvider = getLambdaSheetProvider();

		SkinnyWorkbookProvider workbookProvider = new SkinnyWorkbookProvider(List.of(lambdaSheetProvider));

		File outputFile = writeToFile(targetFolder, workbookProvider.getWorkbook());

		actualWorkbook = new XSSFWorkbook(outputFile);
	}

	private static SheetProvider getLambdaSheetProvider() {
		ColumnHeaderSupplier columnHeaderSupplier =
				() -> List.of("Header-0", "Header-1", "Header-2", "Header-3", "Header-4");
		SheetContentSupplier sheetContentSupplier = () -> List.of(
				() -> List.of("Value-0-0-0", "Value-0-0-1", "Value-0-0-2", "Value-0-0-3", "Value-0-0-4"),
				() -> List.of("Value-0-1-0", "Value-0-1-1", "Value-0-1-2", "Value-0-1-3", "Value-0-1-4"),
				() -> List.of("Value-0-2-0", "Value-0-2-1", "Value-0-2-2", "Value-0-2-3", "Value-0-2-4", "Value-0-2-5"),
				() -> List.of("Value-0-3-0", "Value-0-3-1", "Value-0-3-2", "Value-0-3-3"));

		return new SkinnySheetProvider(sheetContentSupplier, LAMBDA_SHEET_NAME, columnHeaderSupplier);
	}

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

}
