package nl.neutius.xlsx.skinny.writer;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

class SkinnyWriterTest extends AbstractSkinnyWriterTestBase {

    @Test
    void addSeveralRows_allRowsArePresent(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();
        writer.addRowToCurrentSheet(List.of("entry0"));
        writer.addRowToCurrentSheet(List.of("entry1"));
        writer.addRowToCurrentSheet(List.of("entry2"));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
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

        writer.createNewXlsxFile();
        writer.addSeveralRowsToCurrentSheet(List.of(List.of("entry0"), List.of("entry1"), List.of("entry2")));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
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

        writer.createNewXlsxFile();
        writer.addSeveralRowsToCurrentSheet(List.of(firstRow, secondRow, thirdRow));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);

        assertThat(actualSheet.getRow(0)).isNotNull().isNotEmpty().hasSize(5);
        verifyCellContent(actualSheet, 0, 0, "entry0");
        verifyCellContent(actualSheet, 0, 1, "1");
        verifyCellContent(actualSheet, 0, 2, "?");
        verifyCellContent(actualSheet, 0, 3, "Mariënberg");
        verifyCellContent(actualSheet, 0, 4, "Curaçao");

        assertThat(actualSheet.getRow(1)).isNotNull().isNotEmpty().hasSize(4);
        verifyCellContent(actualSheet, 1, 1, "false");
        verifyCellContent(actualSheet, 1, 2, "true");
        verifyCellContent(actualSheet, 1, 3, "null");

        assertThat(actualSheet.getRow(2)).isNotNull().isNotEmpty().hasSize(6);
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

        writer.createNewXlsxFile();
        writer.addRowToCurrentSheet(List.of(valueWithTabs, valueWithNewLines, valueWithSpecialCharacters));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(1);
        assertThat(actualSheet.getRow(1)).isNull();

        Row actualRow = actualSheet.getRow(0);
        assertThat(actualRow).isNotNull().isNotEmpty().hasSize(3);
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

        writer.createNewXlsxFile();
        writer.addRowToCurrentSheet(entryList);
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotNull().isNotEmpty().hasSize(1);

        verifyCellContent(actualSheet, 0, 0, "value");
        verifyCellContent(actualSheet, 0, 1, "");
        verifyCellContent(actualSheet, 0, 2, "");
        verifyCellContent(actualSheet, 0, 3, "");
        verifyCellContent(actualSheet, 0, 4, "");
        verifyCellContent(actualSheet, 0, 5, "value2");
    }

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

        writer.createNewXlsxFile();
        writer.addSeveralRowsToCurrentSheet(List.of(firstRow, secondRow, thirdRow));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
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

}