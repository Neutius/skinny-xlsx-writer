package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterSheetHandlingTest extends AbstractSkinnyWriterTestBase {

    @Test
    void valuesWithDifferentLengthAndHeight_columnWidthIsAdjusted(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String longValue = "we like writing an entire book in a single cell within a bigger spreadsheet, " +
                "also, we \t like \t tabs \t\t\t too \t much";
        String valueWithSpecialCharacters = " \\ \\ \" \" ; || ;; | , . \" ";
        String valueWithSeveralNewLines = "we \n use \n new \n lines \n\n\n\n within a \n single \n value";
        String valueWithSingleNewLines = "First sentence.\nSecond sentence.";

        List<String> firstRow = List.of("short", "value", valueWithSpecialCharacters, longValue);
        List<String> secondRow = List.of("short", valueWithSeveralNewLines, "value");
        List<String> thirdRow = List.of("short", valueWithSingleNewLines, "value");
        writer.addSeveralRowsToCurrentSheet(List.of(firstRow, secondRow, thirdRow));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(3);

        int firstColumnWidth = actualSheet.getColumnWidth(0);
        int secondColumnWidth = actualSheet.getColumnWidth(1);
        int thirdColumnWidth = actualSheet.getColumnWidth(2);
        int fourthColumnWidth = actualSheet.getColumnWidth(3);

        assertThat(firstColumnWidth).isNotEqualTo(secondColumnWidth);
        assertThat(secondColumnWidth).isNotEqualTo(thirdColumnWidth);
        assertThat(thirdColumnWidth).isNotEqualTo(fourthColumnWidth);

        assertThat(secondColumnWidth).isBetween(firstColumnWidth, thirdColumnWidth);
        assertThat(thirdColumnWidth).isBetween(secondColumnWidth, fourthColumnWidth);
    }

    @Test
    void addSheetToWorkbook_fileHasTwoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook("second sheet");

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(2);
        assertThat(actualWorkbook.getSheetAt(0)).isNotNull().hasSize(0);
        assertThat(actualWorkbook.getSheetAt(1)).isNotNull().hasSize(0);
        assertThatThrownBy(() -> actualWorkbook.getSheetAt(2)).isInstanceOf(IllegalArgumentException.class);

        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isEqualTo("second sheet");
    }

    @Test
    void addSeveralSheets_contentHasVaryingLength_columnWidthAdjustedForAllSheetsAndColumns(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addRowToCurrentSheet(List.of("1", "123", "1234567", "123456789"));
        writer.addSheetToWorkbook("second sheet");
        writer.addRowToCurrentSheet(List.of("123456789", "1", "123", "1234567"));
        writer.addSheetToWorkbook("third sheet");
        writer.addRowToCurrentSheet(List.of("123456789", "1234567", "1", "123"));

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).hasSize(3);

        XSSFSheet firstSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(firstSheet).isNotNull().hasSize(1);

        XSSFSheet secondSheet = actualWorkbook.getSheet("second sheet");
        assertThat(secondSheet).isNotNull().hasSize(1);

        XSSFSheet thirdSheet = actualWorkbook.getSheet("third sheet");
        assertThat(thirdSheet).isNotNull().hasSize(1);

        verifySameColumnWidth(firstSheet, 0, secondSheet, 1);
        verifySameColumnWidth(firstSheet, 0, thirdSheet, 2);
        verifySameColumnWidth(firstSheet, 1, secondSheet, 2);
        verifySameColumnWidth(firstSheet, 1, thirdSheet, 3);
        verifySameColumnWidth(firstSheet, 2, secondSheet, 3);
        verifySameColumnWidth(firstSheet, 2, thirdSheet, 1);
        verifySameColumnWidth(firstSheet, 3, secondSheet, 0);
        verifySameColumnWidth(firstSheet, 3, thirdSheet, 0);
    }

    private void verifySameColumnWidth(XSSFSheet sheet, int columnIndex, XSSFSheet otherSheet, int otherColumnIndex) {
        assertThat(sheet.getColumnWidth(columnIndex)).isEqualTo(otherSheet.getColumnWidth(otherColumnIndex));
    }

    @Test
    void addSheetWithEmptyStringAsName_nameIsGenerated(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook("");

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(2);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isEqualTo("Sheet_2");
    }

    @Test
    void addSheetWithNullAsName_nameIsGenerated(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook((String) null);

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(2);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isEqualTo("Sheet_2");
    }

    @Test
    void addSheetWithTooLongName_nameIsSnipped(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook("abcdefghijklmnopqrstuvwxyz1234567890");

        writeAndReadActualWorkbook(targetFolder);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(2);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isEqualTo("abcdefghijklmnopqrstuvwxyz12345");
    }

}