package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterInterfaceHandlingTest extends AbstractSkinnyWriterTestBase {

    private final List<List<String>> contentRows = List.of(List.of("A1", "A2"), List.of("B1", "B2"));
    private final List<String> columnHeaders = List.of("Header1", "Header2");

    private SkinnySheetContent firstSheetContent;
    private XSSFSheet actualSheet;

    @Test
    void addSheetWithNameAndContent_fileHasSheet(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setupWithoutHeaders(targetFolder);

        assertThat(actualWorkbook).hasSize(1);
    }

    @Test
    void addSheetWithNameAndContent_sheetHasCorrectSize(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setupWithoutHeaders(targetFolder);

        assertThat(actualSheet).hasSize(2);
        assertThat(actualSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(2);
    }

    @Test
    void addSheetWithNameAndContent_sheetHasCorrectContent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setupWithoutHeaders(targetFolder);

        verifyCellContent(actualSheet, 0, 0, "A1");
        verifyCellContent(actualSheet, 0, 1, "A2");
        verifyCellContent(actualSheet, 1, 0, "B1");
        verifyCellContent(actualSheet, 1, 1, "B2");
    }

    @Test
    void addSheetWithNameAndHeadersAndContent_sheetHasCorrectSize(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setUpWithHeaders(targetFolder);

        Assertions.assertThat(actualSheet).hasSize(3);
        assertThat(actualSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(2);
    }

    @Test
    void addSheetWithNameAndHeadersAndContent_sheetHasCorrectContent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setUpWithHeaders(targetFolder);

        verifyCellContent(actualSheet, 1, 0, "A1");
        verifyCellContent(actualSheet, 1, 1, "A2");
        verifyCellContent(actualSheet, 2, 0, "B1");
        verifyCellContent(actualSheet, 2, 1, "B2");
    }

    @Test
    void addSheetWithNameAndHeadersAndContent_sheetHasCorrectHeaders(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        setUpWithHeaders(targetFolder);

        verifyCellContent(actualSheet, 0, 0, "Header1");
        verifyCellContent(actualSheet, 0, 1, "Header2");

        verifyFreezePane(actualSheet);
        verifyColumnHeaderFont(actualSheet);
        verifyContentCellFont(actualSheet);
    }

    private void verifyFreezePane(XSSFSheet sheet) {
        PaneInformation paneInformation = sheet.getPaneInformation();
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

    private void verifyColumnHeaderFont(XSSFSheet sheet) {
        XSSFRow columnHeaderRow = sheet.getRow(0);
        for (Cell cell : columnHeaderRow) {
            XSSFRichTextString headerCellValue = (XSSFRichTextString) cell.getRichStringCellValue();
            assertThat(headerCellValue.getFontAtIndex(0)).isNotNull();
            assertThat(headerCellValue.getFontAtIndex(0).getBold()).isTrue();
        }
    }

    private void verifyContentCellFont(XSSFSheet sheet) {
        XSSFRichTextString contentCellValue = sheet.getRow(1).getCell(0).getRichStringCellValue();
        assertThat(contentCellValue.getFontAtIndex(0)).isNull();
    }

    private void setupWithoutHeaders(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        firstSheetContent = DefaultSheetContent.withoutHeaders(SHEET_NAME, contentRows);
        writeInterfaceToFile(targetFolder);
        actualSheet = actualWorkbook.getSheetAt(0);
    }

    private void setUpWithHeaders(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        firstSheetContent = DefaultSheetContent.withHeaders(SHEET_NAME, columnHeaders, contentRows);
        writeInterfaceToFile(targetFolder);
        actualSheet = actualWorkbook.getSheetAt(0);
    }

    private void writeInterfaceToFile(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME);
        writer.addSheetToWorkbook(firstSheetContent);
        writeAndReadActualWorkbook(targetFolder);
    }

}