package nl.neutius.xlsx.skinny.writer;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

class SkinnyWriterTest extends AbstractSkinnyWriterTestBase {

    protected File targetFolder;

    @BeforeEach
    void setUpTargetFolder() {
        targetFolder = new File("D:\\dev\\test\\excel");
        File[] fileArray = targetFolder.listFiles();

        if (fileArray != null && fileArray.length > 0) {
            for (File file : fileArray) {
                assertThat(file.delete()).isTrue();
            }
        }
    }

    @Test
    void verifySetUp_targetFolderIsUsable() {
        assertThat(targetFolder).exists();
        assertThat(targetFolder).isDirectory();
        assertThat(targetFolder).isEmptyDirectory();
        assertThat(targetFolder).canRead();
        assertThat(targetFolder).canWrite();
    }

    @Test
    void createNewFile_fileExists() throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        assertThat(expectedFile).exists();
    }

    @Test
    void createNewFile_emptyFileIsValidXlsxFile() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File targetFile = new File(targetFolder, FILE_NAME + EXTENSION);
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(targetFile);
        XSSFSheet actualSheet = actualWorkbook.getSheetAt(0);

        assertThat(actualSheet).isNotNull();
        assertThat(actualSheet).isEmpty();
        assertThat(actualSheet).hasSize(0);
        assertThat(actualSheet.getActiveCell()).isNull();
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(-1);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(-1);
        assertThat(actualSheet.getCellComments()).isEmpty();
        assertThat(actualSheet.getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualSheet).isEqualTo(actualWorkbook.getSheet(SHEET_NAME));
    }

    @Test
    void addContent_fileHasContent() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();
        writer.addRowToCurrentSheet(List.of("entry"));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotEmpty();
        assertThat(actualSheet).hasSize(1);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(0);

        Row actualRow = actualSheet.getRow(0);
        assertThat(actualRow).isNotNull();
        assertThat(actualRow).isNotEmpty();
        assertThat(actualRow).hasSize(1);
        assertThat((int) actualRow.getFirstCellNum()).isEqualTo(0);
        assertThat((int) actualRow.getLastCellNum()).isEqualTo(1);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(1);
        assertThat(actualRow.getRowNum()).isEqualTo(0);

        Cell actualCell = actualRow.getCell(0);
        assertThat(actualCell).isNotNull();
        assertThat(actualCell.getSheet()).isEqualTo(actualSheet);
        assertThat(actualCell.getStringCellValue()).isEqualTo("entry");
        assertThat(actualCell.getColumnIndex()).isEqualTo(0);
        assertThat(actualCell.getRowIndex()).isEqualTo(0);
    }

    @Test
    void addSeveralRows_allRowsArePresent() throws IOException, InvalidFormatException {
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

        assertThat(actualSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("entry0");
        assertThat(actualSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("entry1");
        assertThat(actualSheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("entry2");
    }

    @Test
    void addSeveralRowsAsOneList_allRowsArePresent() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();
        writer.addSeveralRowsToCurrentSheet(List.of(List.of("entry0"), List.of("entry1"), List.of("entry2")));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).hasSize(3);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(2);

        assertThat(actualSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("entry0");
        assertThat(actualSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("entry1");
        assertThat(actualSheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("entry2");
    }

    @Test
    void addSeveralRowsAndColumns_allFieldsArePresent() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        List<String> firstRow = List.of("entry0", "1", "?", "Mariënberg", "Curaçao");
        List<String> secondRow = List.of("entry1", "false", "true", "null");
        List<String> thirdRow = List.of("entry2", "", "", "", "", "sixth      column");

        writer.createNewXlsxFile();
        writer.addSeveralRowsToCurrentSheet(List.of(firstRow, secondRow, thirdRow));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);

        Row actualFirstRow = actualSheet.getRow(0);
        assertThat(actualFirstRow).isNotNull().isNotEmpty().hasSize(5);
        assertThat(actualFirstRow.getCell(0).getStringCellValue()).isEqualTo("entry0");
        assertThat(actualFirstRow.getCell(1).getStringCellValue()).isEqualTo("1");
        assertThat(actualFirstRow.getCell(2).getStringCellValue()).isEqualTo("?");
        assertThat(actualFirstRow.getCell(3).getStringCellValue()).isEqualTo("Mariënberg");
        assertThat(actualFirstRow.getCell(4).getStringCellValue()).isEqualTo("Curaçao");

        Row actualSecondRow = actualSheet.getRow(1);
        assertThat(actualSecondRow).isNotNull().isNotEmpty().hasSize(4);
        assertThat(actualSecondRow.getCell(1).getStringCellValue()).isEqualTo("false");
        assertThat(actualSecondRow.getCell(2).getStringCellValue()).isEqualTo("true");
        assertThat(actualSecondRow.getCell(3).getStringCellValue()).isEqualTo("null");

        Row actualThirdRow = actualSheet.getRow(2);
        assertThat(actualThirdRow).isNotNull().isNotEmpty().hasSize(6);
        assertThat(actualThirdRow.getCell(0).getStringCellValue()).isEqualTo("entry2");
        assertThat(actualThirdRow.getCell(1).getStringCellValue()).isEqualTo("");
        assertThat(actualThirdRow.getCell(2).getStringCellValue()).isEqualTo("");
        assertThat(actualThirdRow.getCell(3).getStringCellValue()).isEqualTo("");
        assertThat(actualThirdRow.getCell(4).getStringCellValue()).isEqualTo("");
        assertThat(actualThirdRow.getCell(5).getStringCellValue()).isEqualTo("sixth      column");
    }

    @Test
    void addValuesWithWhiteSpaceCharacters_contentHasSameWhiteSpaceCharacters() throws IOException, InvalidFormatException {
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
    void addRowWithNullValues_areConvertedToEmptyStrings() throws IOException, InvalidFormatException {
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

        Row actualRow = actualSheet.getRow(0);
        assertThat(actualRow.getCell(0).getStringCellValue()).isEqualTo("value");
        assertThat(actualRow.getCell(1).getStringCellValue()).isEqualTo("");
        assertThat(actualRow.getCell(2).getStringCellValue()).isEqualTo("");
        assertThat(actualRow.getCell(3).getStringCellValue()).isEqualTo("");
        assertThat(actualRow.getCell(4).getStringCellValue()).isEqualTo("");
        assertThat(actualRow.getCell(5).getStringCellValue()).isEqualTo("value2");
    }

    @Test
    void valuesWithDifferentLengthAndHeight_columnWidthIsAdjusted() throws IOException, InvalidFormatException {
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