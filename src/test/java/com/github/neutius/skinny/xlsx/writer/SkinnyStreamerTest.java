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
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyStreamerTest extends AbstractSkinnyWriterTestBase {

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
    void writeContentToFileSystem_contentCellsAreNotBold(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withHeaders(SHEET_NAME, List.of("Header 1", "Header 2"),
                List.of(List.of("Content 1", "Content 2")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);

        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(2);

        XSSFRow row = actualSheet.getRow(1);
        for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
            XSSFFont font = row.getCell(index).getCellStyle().getFont();
            assertThat(font).isNotNull();
            assertThat(font.getBold()).isFalse();
        }
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

        XSSFRow row = actualSheet.getRow(0);
        for (int index = 0; index < row.getPhysicalNumberOfCells(); index++) {
            XSSFFont font = row.getCell(index).getCellStyle().getFont();
            assertThat(font).isNotNull();
            assertThat(font.getBold()).isTrue();
        }
    }

    @Test
    void writeContentToFileSystem_firstSheetHasFreezePane(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        SkinnySheetContent firstSheet = DefaultSheetContent.withHeaders(SHEET_NAME, List.of("Header 1", "Header 2"),
                List.of(List.of("Content 1", "Content 2")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(firstSheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);

        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(2);

        PaneInformation paneInformation = actualSheet.getPaneInformation();
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

    @Test
    void writeContentToFileSystem_withAutoSizeColumn_withoutHeaders_columnsHaveDifferentSize(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        SkinnySheetContent sheet = DefaultSheetContent.withoutHeaders(SHEET_NAME,
                List.of(List.of("Short", "Medium-sized text", "Longer text to be added to content cell")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(sheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
    }

    @Test
    void writeContentToFileSystem_noAutoSizeColumn_withoutHeaders_columnsHaveSameSize(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        SkinnySheetContent sheet = DefaultSheetContent.withoutHeaders(SHEET_NAME,
                List.of(List.of("Short", "Medium-sized text", "Longer text to be added to content cell")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(sheet), false);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet.getColumnWidth(0)).isEqualTo(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isEqualTo(actualSheet.getColumnWidth(2));
    }

    @Test
    void writeContentToFileSystem_withAutoSizeColumn_withHeaders_columnsHaveDifferentSize(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        SkinnySheetContent sheet = DefaultSheetContent.withHeaders(SHEET_NAME,
                List.of("Short", "Medium-sized text", "Longer text to be added to content cell"),
                List.of(List.of("")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(sheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
    }

    @Test
    void writeContentToFileSystem_noAutoSizeColumn_withHeaders_columnsHaveSameSize(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        SkinnySheetContent sheet = DefaultSheetContent.withHeaders(SHEET_NAME,
                List.of("Short", "Medium-sized text", "Longer text to be added to content cell"),
                List.of(List.of("")));

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(sheet), false);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet.getColumnWidth(0)).isEqualTo(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isEqualTo(actualSheet.getColumnWidth(2));
    }

    @Test
    void writeContentToFileSystem_nullValuePassedIn_emptyRowIsAdded(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        List<List<String>> contentRows = new ArrayList<>();
        contentRows.add(List.of("Cell Content", "More Content"));
        contentRows.add(null);
        contentRows.add(List.of("Row 3 Cell 1", "Row 3 Cell 2"));
        SkinnySheetContent sheet = DefaultSheetContent.withoutHeaders(SHEET_NAME, contentRows);

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, List.of(sheet), true);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);

        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(3);
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