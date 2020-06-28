package nl.neutius.xlsx.skinny.writer;

import static org.assertj.core.api.Assertions.assertThat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.List;

class SkinnyWriterTest extends AbstractSkinnyWriterTestBase {

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
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

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