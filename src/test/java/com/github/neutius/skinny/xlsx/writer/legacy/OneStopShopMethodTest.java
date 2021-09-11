package com.github.neutius.skinny.xlsx.writer.legacy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class OneStopShopMethodTest extends SkinnyWriterInterfaceHandlingTest {

    @Test
    void addSeveralSheetsWithAndWithoutHeaders_allSheetsComeOutProperly(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        List<SkinnySheetContent> sheetContentList = getSheetContentList();

        writer = new SkinnyWriter(targetFolder, FILE_NAME);
        writer.addSeveralSheetsToWorkbook(sheetContentList);
        writer.writeToFile();

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).hasSize(4);
        verifyFirstSheet();
        verifySecondSheet();
        verifyThirdSheet();
        verifyFourthSheet();
    }

    @Test
    void useSkinnyWriterStaticFactoryMethod_allSheetsComeOutProperly(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        List<SkinnySheetContent> sheetContentList = getSheetContentList();

        SkinnyWriter.writeContentToFileSystem(targetFolder, FILE_NAME, sheetContentList);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).hasSize(4);
        verifyFirstSheet();
        verifySecondSheet();
        verifyThirdSheet();
        verifyFourthSheet();
    }

    @Test
    void useSkinnyStreamerStaticFactoryMethod_allSheetsComeOutProperly(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        List<SkinnySheetContent> sheetContentList = getSheetContentList();

        SkinnyStreamer.writeContentToFileSystem(targetFolder, FILE_NAME, sheetContentList);

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        assertThat(actualWorkbook).hasSize(4);
        verifyFirstSheet();
        verifySecondSheet();
        verifyThirdSheet();
        verifyFourthSheet();
    }

    private List<SkinnySheetContent> getSheetContentList() {
        firstSheetContent = DefaultSheetContent.withHeaders(SHEET_NAME, columnHeaders, contentRows);

        List<List<String>> secondSheetContent = List.of(List.of("XX", "XY"), List.of("YX", "YY"));
        SkinnySheetContent secondSheet = DefaultSheetContent.withoutHeaders("second sheet", secondSheetContent);

        SkinnySheetContent thirdSheet = DefaultSheetContent.withHeaders(SHEET_NAME, columnHeaders, Collections.emptyList());

        List<List<String>> fourthSheetContent = List.of(List.of("1", "2", "3", "4"));
        SkinnySheetContent fourthSheet = DefaultSheetContent.withoutHeaders("sheet4", fourthSheetContent);

        return List.of(firstSheetContent, secondSheet, thirdSheet, fourthSheet);
    }

    private void verifyFirstSheet() {
        XSSFSheet actualFirstSheet = actualWorkbook.getSheetAt(0);
        assertThat(actualFirstSheet.getSheetName()).isEqualTo(SHEET_NAME);
        verifyColumnHeaders(actualFirstSheet);

        assertThat(actualFirstSheet).hasSize(3);
        assertThat(actualFirstSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualFirstSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualFirstSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(2);

        verifyCellContent(actualFirstSheet, 0, 0, "Header1");
        verifyCellContent(actualFirstSheet, 0, 1, "Header2");
        verifyCellContent(actualFirstSheet, 1, 0, "A1");
        verifyCellContent(actualFirstSheet, 1, 1, "A2");
        verifyCellContent(actualFirstSheet, 2, 0, "B1");
        verifyCellContent(actualFirstSheet, 2, 1, "B2");
    }

    private void verifySecondSheet() {
        XSSFSheet actualSecondSheet = actualWorkbook.getSheetAt(1);
        assertThat(actualSecondSheet.getSheetName()).isEqualTo("second sheet");
        verifyNoColumnHeaders(actualSecondSheet);

        assertThat(actualSecondSheet).hasSize(2);
        assertThat(actualSecondSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);
        assertThat(actualSecondSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(2);

        verifyCellContent(actualSecondSheet, 0, 0, "XX");
        verifyCellContent(actualSecondSheet, 0, 1, "XY");
        verifyCellContent(actualSecondSheet, 1, 0, "YX");
        verifyCellContent(actualSecondSheet, 1, 1, "YY");
    }

    private void verifyThirdSheet() {
        XSSFSheet actualThirdSheet = actualWorkbook.getSheetAt(2);
        assertThat(actualThirdSheet.getSheetName()).isNotNull().isNotBlank().isNotEqualTo(SHEET_NAME);
        verifyFreezePane(actualThirdSheet);
        verifyColumnHeaderFont(actualThirdSheet);

        assertThat(actualThirdSheet).hasSize(1);
        assertThat(actualThirdSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(2);

        verifyCellContent(actualThirdSheet, 0, 0, "Header1");
        verifyCellContent(actualThirdSheet, 0, 1, "Header2");
    }

    private void verifyFourthSheet() {
        XSSFSheet actualFourthSheet = actualWorkbook.getSheetAt(3);
        assertThat(actualFourthSheet.getSheetName()).isEqualTo("sheet4");
        verifyNoColumnHeaders(actualFourthSheet);

        assertThat(actualFourthSheet).hasSize(1);
        assertThat(actualFourthSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(4);

        verifyCellContent(actualFourthSheet, 0, 0, "1");
        verifyCellContent(actualFourthSheet, 0, 1, "2");
        verifyCellContent(actualFourthSheet, 0, 2, "3");
        verifyCellContent(actualFourthSheet, 0, 3, "4");
    }

}