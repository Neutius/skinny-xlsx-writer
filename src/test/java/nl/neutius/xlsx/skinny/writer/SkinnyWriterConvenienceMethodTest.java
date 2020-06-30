package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterConvenienceMethodTest extends AbstractSkinnyWriterTestBase {

    @Test
    void addSheetAndContentToWorkbook_fileHasTwoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String sheetName = "second sheet";
        List<List<String>> sheetContent = List.of(List.of("11", "12"), List.of("21", "22"));
        writer.addSheetWithContentToWorkbook(sheetName, sheetContent);
        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

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
        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

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

        XSSFCell contentCell = secondSheet.getRow(1).getCell(0);
        assertThat(contentCell.getRichStringCellValue().getFontAtIndex(0)).isNull();
        assertThat(contentCell.getCellStyle().getWrapText()).isTrue();

        XSSFCell headerCell = secondSheet.getRow(0).getCell(0);
        assertThat(headerCell.getRichStringCellValue().getFontAtIndex(0).getBold()).isTrue();
        assertThat(headerCell.getCellStyle().getWrapText()).isFalse();

        PaneInformation paneInformation = secondSheet.getPaneInformation();
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

    @Test
    void addSeveralSheetsWithContentToWorkbook_fileHasTwoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        String secondSheetName = "second sheet";
        List<List<String>> secondSheetContent = List.of(List.of("11", "12"), List.of("21", "22"));
        String thirdSheetName = "third sheet";
        List<List<String>> thirdSheetContent = List.of(List.of("A", "B"), List.of("C", "D"));

        Map<String, List<List<String>>> sheetNameAndContentMap = new HashMap<>();
        sheetNameAndContentMap.put(secondSheetName, secondSheetContent);
        sheetNameAndContentMap.put(thirdSheetName, thirdSheetContent);

        writer.addSeveralSheetsWithContentToWorkbook(sheetNameAndContentMap);
        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

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

}