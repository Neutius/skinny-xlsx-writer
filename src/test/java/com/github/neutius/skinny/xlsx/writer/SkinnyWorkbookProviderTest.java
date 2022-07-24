package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWorkbookProviderTest {
    private static final String VALUE_1 = "value-1";
    private static final String SHEET_NAME = "sheet name";

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

    @Test
    void useLambdasForSingleValue_valueIsPresent() {
        SkinnySheetProvider provider = new SkinnySheetProvider(() -> List.of(() -> List.of("value-1")));
        testSubject = new SkinnyWorkbookProvider(List.of(provider));

        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("value-1");
    }

    @Test
    void addSeveralSheetsWithNoName_sheetsHaveUniqueNames() {
        SheetContentSupplier sameSheet = () -> List.of(() -> List.of(VALUE_1));
        SkinnySheetProvider provider = new SkinnySheetProvider(sameSheet);
        testSubject = new SkinnyWorkbookProvider(List.of(provider, provider, provider));

        Workbook workbook = testSubject.getWorkbook();
        String sheetName0 = workbook.getSheetAt(0).getSheetName();
        String sheetName1 = workbook.getSheetAt(1).getSheetName();
        String sheetName2 = workbook.getSheetAt(2).getSheetName();

        assertThat(sheetName0).isNotEqualTo(sheetName1);
        assertThat(sheetName0).isNotEqualTo(sheetName2);
        assertThat(sheetName1).isNotEqualTo(sheetName2);
    }

    @Test
    void addSheetWithName_sheetHasName() {
        SheetContentSupplier contentSupplier = () -> List.of(() -> List.of(VALUE_1));
        SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME);
        testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

        Workbook workbook = testSubject.getWorkbook();

        assertThat(workbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
    }

    @Test
    void addSeveralSheetsWithTheSameName_sheetsHaveUniqueNames() {
        SheetContentSupplier sameSheet = () -> List.of(() -> List.of(VALUE_1));
        SkinnySheetProvider provider = new SkinnySheetProvider(sameSheet, SHEET_NAME);
        testSubject = new SkinnyWorkbookProvider(List.of(provider, provider, provider));

        Workbook workbook = testSubject.getWorkbook();
        String sheetName0 = workbook.getSheetAt(0).getSheetName();
        String sheetName1 = workbook.getSheetAt(1).getSheetName();
        String sheetName2 = workbook.getSheetAt(2).getSheetName();

        assertThat(sheetName0).isNotEqualTo(sheetName1);
        assertThat(sheetName0).isNotEqualTo(sheetName2);
        assertThat(sheetName1).isNotEqualTo(sheetName2);
    }

    @Test
    void cellValuesHaveDifferentWidth_sheetColumnsHaveDifferentWidth() {
        SheetContentSupplier sheet = () -> List.of(() -> List.of("short", "medium sized", "relatively large piece of text"));
        testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(sheet, SHEET_NAME)));

        Workbook workbook = testSubject.getWorkbook();
        Sheet actualSheet = workbook.getSheetAt(0);

        assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
    }

    @Test
    void cellValuesHaveDifferentWidthOnlyAfter100Rows_sheetColumnsHaveSameWidth() {
        List<RowContentSupplier> rowContentSupplierList = new ArrayList<>();
        for (int i = 1; i <= 100; i++) {
            rowContentSupplierList.add(Collections::emptyList);
        }
        rowContentSupplierList.add(() -> List.of("short", "medium sized", "relatively large piece of text"));
        testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(() -> rowContentSupplierList, SHEET_NAME)));

        Workbook workbook = testSubject.getWorkbook();
        Sheet actualSheet = workbook.getSheetAt(0);

        assertThat(actualSheet.getColumnWidth(0)).isEqualTo(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isEqualTo(actualSheet.getColumnWidth(2));
    }

    @Test
    void cellValuesHaveDifferentWidthOnlyAfter99Rows_sheetColumnsHaveDifferentWidth() {
        List<RowContentSupplier> rowContentSupplierList = new ArrayList<>();
        for (int i = 1; i <= 99; i++) {
            rowContentSupplierList.add(Collections::emptyList);
        }
        rowContentSupplierList.add(() -> List.of("short", "medium sized", "relatively large piece of text"));
        rowContentSupplierList.add(() -> List.of("short", "medium sized", "relatively large piece of text"));
        testSubject = new SkinnyWorkbookProvider(List.of(new SkinnySheetProvider(() -> rowContentSupplierList, SHEET_NAME)));

        Workbook workbook = testSubject.getWorkbook();
        Sheet actualSheet = workbook.getSheetAt(0);

        assertThat(actualSheet.getColumnWidth(0)).isLessThan(actualSheet.getColumnWidth(1));
        assertThat(actualSheet.getColumnWidth(1)).isLessThan(actualSheet.getColumnWidth(2));
    }

}
