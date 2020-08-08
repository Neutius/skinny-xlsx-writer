package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterColumnHeadersTest extends AbstractSkinnyWriterTestBase {

    @Test
    void columnHeaderRowIsAdded_firstRowOfSheetIsLocked(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));

        writeAndReadActualWorkbook(targetFolder);
        PaneInformation paneInformation = actualWorkbook.getSheetAt(0).getPaneInformation();
        assertThat(paneInformation).isNotNull();
        assertThat(paneInformation.isFreezePane()).isTrue();
        assertThat((int) paneInformation.getHorizontalSplitTopRow()).isEqualTo(1);
        assertThat((int) paneInformation.getHorizontalSplitPosition()).isEqualTo(1);
    }

    @Test
    void columnHeaderRowIsAdded_contentRowsStartBelowHeaderRow(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheetAt(0);
        assertThat(actualSheet).hasSize(3);
        assertThat(actualSheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("column header");
        assertThat(actualSheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("normal cell content");
        assertThat(actualSheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("normal cell content");
    }

    @Test
    void columnHeaderRowIsAdded_headerCellValueIsBold(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheetAt(0);
        XSSFRichTextString headerCellValue = actualSheet.getRow(0).getCell(0).getRichStringCellValue();
        assertThat(headerCellValue.getFontAtIndex(2)).isNotNull();
        assertThat(headerCellValue.getFontAtIndex(2).getBold()).isTrue();
    }

    @Test
    void columnHeaderRowIsAdded_contentCellHasNoFont(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFRichTextString contentCellValue
                = actualWorkbook.getSheetAt(0).getRow(1).getCell(0).getRichStringCellValue();
        assertThat(contentCellValue.getFontAtIndex(2)).isNull();
    }

    @Test
    void columnHeaderRowIsAdded_wrapTextIsFalseForColumnHeadersAndContentCells(@TempDir File targetFolder)
            throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));

        writeAndReadActualWorkbook(targetFolder);
        XSSFSheet actualSheet = actualWorkbook.getSheetAt(0);
        assertThat(actualSheet.getRow(0).getCell(0).getCellStyle().getWrapText()).isFalse();
        assertThat(actualSheet.getRow(1).getCell(0).getCellStyle().getWrapText()).isFalse();
    }

    @Test
    void contentRowsAreAddedFirst_headerRowIsAdded_throwIllegalStateException(@TempDir File targetFolder) throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addRowToCurrentSheet(List.of("normal cell content", "normal cell content"));
        assertThatThrownBy(() -> writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header")))
                .isInstanceOf(IllegalStateException.class)
                .hasCause(null);
    }

    @Test
    void columnHeaderRowIsAddedTwice_throwIllegalStateException(@TempDir File targetFolder) throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header"));
        assertThatThrownBy(() -> writer.addColumnHeaderRowToCurrentSheet(List.of("column header", "column header")))
                .isInstanceOf(IllegalStateException.class)
                .hasCause(null);
    }

    @Test
    void columnHeaderRowWithNullValue_throwsNullPointerException(@TempDir File targetFolder) throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        assertThatThrownBy(() -> writer.addColumnHeaderRowToCurrentSheet(Arrays.asList(null, null)))
                .isInstanceOf(NullPointerException.class)
                .hasCause(null);
    }

    @Test
    void columnHeaderRowWithEmptyString_throwsIllegalArgumentException(@TempDir File targetFolder) throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        assertThatThrownBy(() -> writer.addColumnHeaderRowToCurrentSheet(List.of("", "")))
                .isInstanceOf(IllegalArgumentException.class)
                .hasCause(null);
    }

}