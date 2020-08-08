package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterConvenienceMethodTest extends AbstractSkinnyWriterTestBase {

    @Test
    void callToStringOverride_containsRelevantInformation(@TempDir File targetFolder) throws IOException {
        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        assertThat(writer.toString()).isNotNull().isNotBlank().contains(FILE_NAME)
                .isEqualTo("SkinnyWriter - target .xlsx file: " + expectedFile.toString() +
                        " - current amount of sheets: 1 - current sheet has 0 rows and 0 columns");
    }

    @Test
    void callToStringOverride_containsUpdatedInformation(@TempDir File targetFolder) throws IOException {
        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);
        writer.addSheetWithContentToWorkbook("Sheet2", List.of(List.of("A", "B", "C"), List.of("1", "2", "3")));

        assertThat(writer.toString()).isEqualTo("SkinnyWriter - target .xlsx file: " + expectedFile.toString() +
                " - current amount of sheets: 2 - current sheet has 2 rows and 3 columns");
    }

    @Test
    void addSheetAndContentToWorkbook_fileHasTwoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String sheetName = "second sheet";
        List<List<String>> sheetContent = List.of(List.of("11", "12"), List.of("21", "22"));
        writer.addSheetWithContentToWorkbook(sheetName, sheetContent);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(2);
        assertThatThrownBy(() -> actualWorkbook.getSheetAt(2)).isInstanceOf(IllegalArgumentException.class);

        XSSFSheet firstSheet = actualWorkbook.getSheetAt(0);
        assertThat(firstSheet).isNotNull().hasSize(0);
        assertThat(firstSheet.getSheetName()).isEqualTo(SHEET_NAME);

        XSSFSheet secondSheet = actualWorkbook.getSheetAt(1);
        assertThat(secondSheet).isNotNull().hasSize(2);
        assertThat(secondSheet.getSheetName()).isEqualTo("second sheet");
        assertThat(secondSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("11");
        assertThat(secondSheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("12");
        assertThat(secondSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("21");
        assertThat(secondSheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("22");
    }

    @Test
    void addSheetAndContentWithHeadersToWorkbook_fileHasTwoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        List<List<String>> sheetContent = List.of(List.of("11", "12"), List.of("21", "22"));
        writer.addSheetWithHeadersAndContentToWorkbook("second sheet", sheetContent);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(2);
        assertThatThrownBy(() -> actualWorkbook.getSheetAt(2)).isInstanceOf(IllegalArgumentException.class);

        XSSFSheet firstSheet = actualWorkbook.getSheetAt(0);
        assertThat(firstSheet).isNotNull().hasSize(0);
        assertThat(firstSheet.getSheetName()).isEqualTo(SHEET_NAME);

        XSSFSheet secondSheet = actualWorkbook.getSheetAt(1);
        assertThat(secondSheet).isNotNull().hasSize(2);
        assertThat(secondSheet.getSheetName()).isEqualTo("second sheet");
        assertThat(secondSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("11");
        assertThat(secondSheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("12");
        assertThat(secondSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("21");
        assertThat(secondSheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("22");

        assertColumnHeaderRowAndContentRow(secondSheet);
    }

    @Test
    void addSeveralSheetsWithContentToWorkbook_fileHasSeveralSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String secondSheetName = "second sheet";
        List<List<String>> secondSheetContent = List.of(List.of("11", "12"), List.of("21", "22"));
        String thirdSheetName = "third sheet";
        List<List<String>> thirdSheetContent = List.of(List.of("A", "B"), List.of("C", "D"));
        Map<String, List<List<String>>> sheetNameAndContentMap = new HashMap<>();
        sheetNameAndContentMap.put(secondSheetName, secondSheetContent);
        sheetNameAndContentMap.put(thirdSheetName, thirdSheetContent);
        writer.addSeveralSheetsWithContentToWorkbook(sheetNameAndContentMap);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(3);

        XSSFSheet firstSheet = actualWorkbook.getSheetAt(0);
        assertThat(firstSheet).isNotNull().hasSize(0);
        assertThat(firstSheet.getSheetName()).isEqualTo(SHEET_NAME);

        XSSFSheet secondSheet = actualWorkbook.getSheetAt(1);
        assertThat(secondSheet).isNotNull().hasSize(2);
        assertThat(secondSheet.getSheetName()).isEqualTo("second sheet");

        XSSFSheet thirdSheet = actualWorkbook.getSheetAt(2);
        assertThat(thirdSheet).isNotNull().hasSize(2);
        assertThat(thirdSheet.getSheetName()).isEqualTo("third sheet");
        assertThat(thirdSheet.getPhysicalNumberOfRows()).isEqualTo(2);
    }

    @Test
    void addSeveralSheetsWithContentInLinkedHashMap_fileHasSameSheetsInSameOrder(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        Map<String, List<List<String>>> sheetNameAndContentMap = new LinkedHashMap<>(4);
        sheetNameAndContentMap.put("second sheet", List.of(List.of("2")));
        sheetNameAndContentMap.put("third sheet", List.of(List.of("3")));
        sheetNameAndContentMap.put("fourth sheet", List.of(List.of("4")));
        sheetNameAndContentMap.put(null, List.of(List.of("null")));
        sheetNameAndContentMap.put("sixth sheet", List.of(List.of("6")));
        sheetNameAndContentMap.put("seventh sheet", null);
        sheetNameAndContentMap.put("eighth sheet", List.of(List.of("8")));
        sheetNameAndContentMap.put("ninth sheet", List.of(List.of("9")));
        writer.addSeveralSheetsWithContentToWorkbook(sheetNameAndContentMap);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(9);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isEqualTo("second sheet");
        assertThat(actualWorkbook.getSheetAt(2).getSheetName()).isEqualTo("third sheet");
        assertThat(actualWorkbook.getSheetAt(3).getSheetName()).isEqualTo("fourth sheet");
        assertThat(actualWorkbook.getSheetAt(4).getSheetName()).isNotNull().isNotBlank().isNotEqualTo("null");
        assertThat(actualWorkbook.getSheetAt(5).getSheetName()).isEqualTo("sixth sheet");
        assertThat(actualWorkbook.getSheetAt(6).getSheetName()).isEqualTo("seventh sheet");
        assertThat(actualWorkbook.getSheetAt(7).getSheetName()).isEqualTo("eighth sheet");
        assertThat(actualWorkbook.getSheetAt(8).getSheetName()).isEqualTo("ninth sheet");
    }

    @Test
    void addSeveralSheetsWithHeadersAndContentToWorkbook_allSheetsPresentWithHeadersAndContent(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        Map<String, List<List<String>>> sheetNameAndHeadersAndContentMap = new LinkedHashMap<>();
        sheetNameAndHeadersAndContentMap.put("second sheet",
                List.of(List.of("Header1", "Header2"), List.of("Content1", "Content2")));
        sheetNameAndHeadersAndContentMap.put("third sheet",
                List.of(List.of("Header1", "Header2"), List.of("Content1", "Content2")));
        sheetNameAndHeadersAndContentMap.put("fourth sheet",
                List.of(List.of("Header1", "Header2"), List.of("Content1", "Content2")));
        sheetNameAndHeadersAndContentMap.put(null,
                List.of(List.of("Header1", "Header2"), List.of("Content1", "Content2")));
        writer.addSeveralSheetsWithHeadersAndContentToWorkbook(sheetNameAndHeadersAndContentMap);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(5);
        assertSheetWithColumnHeaderRowAndContentRow(actualWorkbook.getSheetAt(1), "second sheet", 2);
        assertSheetWithColumnHeaderRowAndContentRow(actualWorkbook.getSheetAt(2), "third sheet", 2);
        assertSheetWithColumnHeaderRowAndContentRow(actualWorkbook.getSheetAt(3), "fourth sheet", 2);

        XSSFSheet fifthSheet = actualWorkbook.getSheetAt(4);
        assertThat(fifthSheet).isNotNull().hasSize(2);
        assertThat(fifthSheet.getSheetName()).isNotNull().isNotBlank().isNotEqualTo("null");
        assertColumnHeaderRowAndContentRow(fifthSheet);
    }

    private void assertSheetWithColumnHeaderRowAndContentRow(XSSFSheet actualSheet, String expectedSheetName, int expectedSheetSize) {
        assertThat(actualSheet).isNotNull().hasSize(expectedSheetSize);
        assertThat(actualSheet.getSheetName()).isEqualTo(expectedSheetName);

        assertColumnHeaderRowAndContentRow(actualSheet);
    }

    private void assertColumnHeaderRowAndContentRow(XSSFSheet actualSheet) {
        XSSFCell contentCell = actualSheet.getRow(1).getCell(0);
        assertThat(contentCell.getRichStringCellValue().getFontAtIndex(0)).isNull();
        assertThat(contentCell.getCellStyle().getWrapText()).isFalse();

        XSSFCell headerCell = actualSheet.getRow(0).getCell(0);
        assertThat(headerCell.getRichStringCellValue().getFontAtIndex(0).getBold()).isTrue();
        assertThat(headerCell.getCellStyle().getWrapText()).isFalse();

        PaneInformation paneInformation = actualSheet.getPaneInformation();
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

}