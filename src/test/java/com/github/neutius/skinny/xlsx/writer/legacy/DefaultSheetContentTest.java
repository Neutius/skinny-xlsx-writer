package com.github.neutius.skinny.xlsx.writer.legacy;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

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

    @Test
    void withHeaders_columnHeadersIsNull_throwsIllegalArgumentException() {
        List<String> invalidColumnHeaders = null;

        verifyThrownException(invalidColumnHeaders);
    }

    @Test
    void withHeaders_columnHeadersIsEmptyList_throwsIllegalArgumentException() {
        List<String> invalidColumnHeaders = Collections.emptyList();

        verifyThrownException(invalidColumnHeaders);
    }

    @Test
    void withHeaders_columnHeadersContainsNull_throwsIllegalArgumentException() {
        List<String> invalidColumnHeaders = Arrays.asList("valid", "header", null);

        verifyThrownException(invalidColumnHeaders);
    }

    @Test
    void withHeaders_columnHeadersContainsEmptyString_throwsIllegalArgumentException() {
        List<String> invalidColumnHeaders = Arrays.asList("valid", "header", "");

        verifyThrownException(invalidColumnHeaders);
    }

    @Test
    void withHeaders_columnHeadersContainsBlankString_throwsIllegalArgumentException() {
        List<String> invalidColumnHeaders = Arrays.asList("valid", "header", "       ");

        verifyThrownException(invalidColumnHeaders);
    }

    private void verifyThrownException(List<String> invalidColumnHeaders) {
        assertThatThrownBy(() -> DefaultSheetContent.withHeaders(sheetName, invalidColumnHeaders, contentRows))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("column")
                .hasMessageContaining("header");
    }
}