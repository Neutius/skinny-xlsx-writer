package com.github.neutius.skinny.xlsx.legacy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterContentWritingTest extends AbstractSkinnyWriterTestBase {

    @Test
    void addSeveralRows_allRowsArePresent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addRowToCurrentSheet(List.of("entry0"));
        writer.addRowToCurrentSheet(List.of("entry1"));
        writer.addRowToCurrentSheet(List.of("entry2"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).hasSize(3);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(2);

        verifyCellContent(actualSheet, 0, 0, "entry0");
        verifyCellContent(actualSheet, 1, 0, "entry1");
        verifyCellContent(actualSheet, 2, 0, "entry2");
    }

    @Test
    void addSeveralRowsAsOneList_allRowsArePresent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSeveralRowsToCurrentSheet(List.of(List.of("entry0"), List.of("entry1"), List.of("entry2")));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).hasSize(3);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(2);

        verifyCellContent(actualSheet, 0, 0, "entry0");
        verifyCellContent(actualSheet, 1, 0, "entry1");
        verifyCellContent(actualSheet, 2, 0, "entry2");
    }

    @Test
    void addSeveralRowsAndColumns_allFieldsArePresent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        List<String> firstRow = List.of("entry0", "1", "?", "Mariënberg", "Curaçao");
        List<String> secondRow = List.of("entry1", "false", "true", "null");
        List<String> thirdRow = List.of("entry2", "", "", "", "", "sixth      column");
        writer.addSeveralRowsToCurrentSheet(List.of(firstRow, secondRow, thirdRow));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);

        assertThat(actualSheet.getRow(0).getPhysicalNumberOfCells()).isEqualTo(5);
        verifyCellContent(actualSheet, 0, 0, "entry0");
        verifyCellContent(actualSheet, 0, 1, "1");
        verifyCellContent(actualSheet, 0, 2, "?");
        verifyCellContent(actualSheet, 0, 3, "Mariënberg");
        verifyCellContent(actualSheet, 0, 4, "Curaçao");

        assertThat(actualSheet.getRow(1).getPhysicalNumberOfCells()).isEqualTo(4);
        verifyCellContent(actualSheet, 1, 0, "entry1");
        verifyCellContent(actualSheet, 1, 1, "false");
        verifyCellContent(actualSheet, 1, 2, "true");
        verifyCellContent(actualSheet, 1, 3, "null");

        assertThat(actualSheet.getRow(2).getPhysicalNumberOfCells()).isEqualTo(6);
        verifyCellContent(actualSheet, 2, 0, "entry2");
        verifyCellContent(actualSheet, 2, 1, "");
        verifyCellContent(actualSheet, 2, 2, "");
        verifyCellContent(actualSheet, 2, 3, "");
        verifyCellContent(actualSheet, 2, 4, "");
        verifyCellContent(actualSheet, 2, 5, "sixth      column");
    }

    @Test
    void addValuesWithWhiteSpaceCharacters_contentHasSameWhiteSpaceCharacters(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String valueWithTabs = "we \t like \t tabs \t\t\t too \t much";
        String valueWithNewLines = "we \n use \n new \n lines \n\n\n\n within a \n single \n value";
        String valueWithSpecialCharacters = " \\ \\ \" \" ; || ;; | , . \" ";
        writer.addRowToCurrentSheet(List.of(valueWithTabs, valueWithNewLines, valueWithSpecialCharacters));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(1);
        assertThatThrownBy(() -> actualSheet.getRow(1).getPhysicalNumberOfCells())
                .isInstanceOf(NullPointerException.class);

        XSSFRow actualRow = actualSheet.getRow(0);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(3);
        assertThat(actualRow.getCell(0).getStringCellValue()).isEqualTo(valueWithTabs);
        assertThat(actualRow.getCell(1).getStringCellValue()).isEqualTo(valueWithNewLines);
        assertThat(actualRow.getCell(2).getStringCellValue()).isEqualTo(valueWithSpecialCharacters);
        assertThat(actualRow.getCell(3)).isNull();
    }

    @Test
    void addRowWithNullValues_areConvertedToEmptyStrings(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        List<String> entryList = new ArrayList<>();
        entryList.add("value");
        entryList.add("");
        entryList.add("");
        entryList.add(null);
        entryList.add(null);
        entryList.add("value2");
        writer.addRowToCurrentSheet(entryList);

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(1);
        verifyCellContent(actualSheet, 0, 0, "value");
        verifyCellContent(actualSheet, 0, 1, "");
        verifyCellContent(actualSheet, 0, 2, "");
        verifyCellContent(actualSheet, 0, 3, "");
        verifyCellContent(actualSheet, 0, 4, "");
        verifyCellContent(actualSheet, 0, 5, "value2");
    }

    @Test
    void addContent_fileHasContent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addRowToCurrentSheet(List.of("entry"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotEmpty();
        assertThat(actualSheet).hasSize(1);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(0);

        XSSFRow actualRow = actualSheet.getRow(0);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(1);
        assertThat((int) actualRow.getFirstCellNum()).isEqualTo(0);
        assertThat((int) actualRow.getLastCellNum()).isEqualTo(1);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(1);
        assertThat(actualRow.getRowNum()).isEqualTo(0);

        XSSFCell actualCell = actualRow.getCell(0);
        assertThat(actualCell).isNotNull();
        assertThat(actualCell.getSheet()).isEqualTo(actualSheet);
        assertThat(actualCell.getStringCellValue()).isEqualTo("entry");
        assertThat(actualCell.getColumnIndex()).isEqualTo(0);
        assertThat(actualCell.getRowIndex()).isEqualTo(0);
    }
}
