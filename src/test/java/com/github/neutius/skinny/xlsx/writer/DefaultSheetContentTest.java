package com.github.neutius.skinny.xlsx.writer;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class DefaultSheetContentTest {

    private final String sheetName = "sheetName";
    private final List<String> columnHeaders = List.of("Header1", "Header2");
    private final List<List<String>> contentRows = List.of(List.of("A1", "A2"), List.of("B1", "B2"));

    @Test
    void withHeaders() {
        SkinnySheetContent sheetContent = DefaultSheetContent.withHeaders(sheetName, columnHeaders, contentRows);

        assertThat(sheetContent.getSheetName()).isEqualTo(sheetName);
        assertThat(sheetContent.hasColumnHeaders()).isTrue();
        assertThat(sheetContent.getColumnHeaders()).isEqualTo(columnHeaders);
        assertThat(sheetContent.getContentRows()).isEqualTo(contentRows);
    }

    @Test
    void withoutHeaders() {
        SkinnySheetContent sheetContent = DefaultSheetContent.withoutHeaders(sheetName, contentRows);

        assertThat(sheetContent.getSheetName()).isEqualTo(sheetName);
        assertThat(sheetContent.hasColumnHeaders()).isFalse();
        assertThat(sheetContent.getColumnHeaders()).isNull();
        assertThat(sheetContent.getContentRows()).isEqualTo(contentRows);
    }
}