package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyStreamerBasicTest extends AbstractSkinnyWriterTestBase {

    @Test
    void writeContentToFileSystem_fileExists(@TempDir File targetFolder) throws IOException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withoutHeaders(SHEET_NAME, List.of(List.of("Cell Content")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        assertThat(expectedFile).exists().isNotNull().isNotEmpty().isFile().canWrite().canRead();
    }

    @Test
    void writeContentToFileSystem_firstSheetHasContent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withoutHeaders(SHEET_NAME, List.of(List.of("Cell Content")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);
        verifySheetWithOneContentCell(actualWorkbook.getSheet(SHEET_NAME));
    }

    @Test
    void writeContentToFileSystem_severalSheetsHaveTheRightContent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withoutHeaders(SHEET_NAME, List.of(List.of("Cell Content")));
        SkinnySheetContent secondSheet = DefaultSheetContent.withoutHeaders("Second Sheet", List.of(
                List.of("Cell Content", "More Content"),
                List.of("Row 2 Cell 1", "Row 2 Cell 2", "Row 2 Cell 3", "Row 2 Cell 4", "Row 2 Cell 5", "Row 2 Cell 6")));
        SkinnySheetContent thirdSheet = DefaultSheetContent.withoutHeaders("Third Sheet", List.of(List.of("Cell Content")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet, secondSheet, thirdSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(3);
        verifySheetWithOneContentCell(actualWorkbook.getSheet(SHEET_NAME));
        verifySheetWithOneContentCell(actualWorkbook.getSheet("Third Sheet"));

        XSSFSheet actualSecondSheet = actualWorkbook.getSheet("Second Sheet");
        assertThat(actualSecondSheet).isNotNull().isNotEmpty().hasSize(2);
        assertThat(actualSecondSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualSecondSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(6);

        verifyCellContent(actualSecondSheet, 0, 0, "Cell Content");
        verifyCellContent(actualSecondSheet, 0, 1, "More Content");
        verifyCellContent(actualSecondSheet, 1, 0, "Row 2 Cell 1");
        verifyCellContent(actualSecondSheet, 1, 1, "Row 2 Cell 2");
        verifyCellContent(actualSecondSheet, 1, 2, "Row 2 Cell 3");
        verifyCellContent(actualSecondSheet, 1, 3, "Row 2 Cell 4");
        verifyCellContent(actualSecondSheet, 1, 4, "Row 2 Cell 5");
        verifyCellContent(actualSecondSheet, 1, 5, "Row 2 Cell 6");
    }

    @Test
    void writeContentToFileSystem_firstSheetHasColumnHeaders(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withHeaders(SHEET_NAME, List.of("Header 1", "Header 2"),
                List.of(List.of("Content 1", "Content 2")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);

        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(2);

        verifyNoHeaderFormatting(actualSheet.getRow(1));
    }

    @Test
    void writeContentToFileSystem_firstSheetHasBoldColumnHeaders(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withHeaders(SHEET_NAME, List.of("Header 1", "Header 2"),
                List.of(List.of("Content 1", "Content 2")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        File actualFile = new File(targetFolder, FILE_NAME + EXTENSION);
        actualWorkbook = new XSSFWorkbook(actualFile);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);

        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(2);

        verifyHeaderFormatting(actualSheet.getRow(0));
    }

    @Test
    void writeContentToFileSystem_firstSheetHasFreezePan(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withHeaders(SHEET_NAME, List.of("Header 1", "Header 2"),
                List.of(List.of("Content 1", "Content 2")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);

        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(2);

        verifyFreezePane(actualSheet.getPaneInformation());
    }

    private void verifyFreezePane(PaneInformation paneInformation) {
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

    private void verifyHeaderFormatting(XSSFRow row) {
        for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
            XSSFFont font = row.getCell(index).getCellStyle().getFont();
            assertThat(font).isNotNull();
            assertThat(font.getBold()).isTrue();
        }
    }

    private void verifyNoHeaderFormatting(XSSFRow row) {
        for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
            XSSFFont font = row.getCell(index).getCellStyle().getFont();
            assertThat(font).isNotNull();
            assertThat(font.getBold()).isFalse();
        }
    }

    private void verifySheetWithOneContentCell(XSSFSheet actualSheet) {
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(1);

        XSSFRow actualRow = actualSheet.getRow(0);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(1);

        XSSFCell actualCell = actualRow.getCell(0);
        assertThat(actualCell).isNotNull();
        assertThat(actualCell.getStringCellValue()).isNotBlank().isEqualTo("Cell Content");
    }

}