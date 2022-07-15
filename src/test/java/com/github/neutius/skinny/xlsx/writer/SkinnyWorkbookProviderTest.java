package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWorkbookProviderTest {
    private SkinnyWorkbookProvider testSubject;

    @Test
    void createInstance_isEmpty() {
        testSubject = new SkinnyWorkbookProvider();

        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook).isNotNull().isEmpty();
    }

    @Test
    void createInstance_isXlsxWorkbook() {
        testSubject = new SkinnyWorkbookProvider();

        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook).isNotInstanceOf(HSSFWorkbook.class);
    }

    @Test
    void addSheetToEmptyWorkbook_workbookHasSheet() {
        testSubject = new SkinnyWorkbookProvider();
        SkinnySheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier();
        sheetContentSupplier.addContentRow("value1");

        testSubject.addSheet(sheetContentSupplier);
        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook).isNotEmpty();
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("value1");
    }

    @Test
    void addSheetWithSeveralRowsAndColumns_allCellValuesAreInTheRightPlace() {
        testSubject = new SkinnyWorkbookProvider();
        SkinnySheetContentSupplier sheetContentSupplier = new SkinnySheetContentSupplier();
        sheetContentSupplier.addContentRow("0-0-0", "0-0-1", "0-0-2", "0-0-3");
        sheetContentSupplier.addContentRow("0-1-0", "0-1-1", "0-1-2", "0-1-3");
        sheetContentSupplier.addContentRow("0-2-0", "0-2-1", "0-2-2", "0-2-3");
        sheetContentSupplier.addContentRow("0-3-0", "0-3-1", "0-3-2", "0-3-3");

        testSubject.addSheet(sheetContentSupplier);
        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("0-0-0");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue()).isEqualTo("0-0-1");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue()).isEqualTo("0-0-2");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue()).isEqualTo("0-0-3");

        assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo("0-1-0");
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue()).isEqualTo("0-1-1");
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(2).getStringCellValue()).isEqualTo("0-1-2");
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue()).isEqualTo("0-1-3");

        assertThat(workbook.getSheetAt(0).getRow(2).getCell(0).getStringCellValue()).isEqualTo("0-2-0");
        assertThat(workbook.getSheetAt(0).getRow(2).getCell(1).getStringCellValue()).isEqualTo("0-2-1");
        assertThat(workbook.getSheetAt(0).getRow(2).getCell(2).getStringCellValue()).isEqualTo("0-2-2");
        assertThat(workbook.getSheetAt(0).getRow(2).getCell(3).getStringCellValue()).isEqualTo("0-2-3");

        assertThat(workbook.getSheetAt(0).getRow(3).getCell(0).getStringCellValue()).isEqualTo("0-3-0");
        assertThat(workbook.getSheetAt(0).getRow(3).getCell(1).getStringCellValue()).isEqualTo("0-3-1");
        assertThat(workbook.getSheetAt(0).getRow(3).getCell(2).getStringCellValue()).isEqualTo("0-3-2");
        assertThat(workbook.getSheetAt(0).getRow(3).getCell(3).getStringCellValue()).isEqualTo("0-3-3");
    }

    @Disabled("TODO write test -> adjust implementation if needed - GvdNL 15-07-2022")
    @Test
    void addSeveralSheetsWithNoName_sheetsHaveUniqueNames() {
        // TODO write test -> adjust implementation if needed - GvdNL 15-07-2022
    }

}
