package com.github.neutius.skinny.xlsx.writer;

import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyRowContentProviderTest {
    private static final String VALUE_1 = "value1";
    private static final String VALUE_2 = "value2";
    private static final String VALUE_3 = "value3";

    private SkinnyRowContentProvider testSubject;

    @Test
    void createInstance_getRowContent_rowContentIsEmpty() {
        testSubject = new SkinnyRowContentProvider();

        List<String> rowContent = testSubject.getRowContent();

        assertThat(rowContent).isNotNull().isEmpty();
    }

    @Test
    void addCellContent_getRowContent_cellContentIsIncluded() {
        testSubject = new SkinnyRowContentProvider();
        testSubject.addCellContent(VALUE_1);

        List<String> rowContent = testSubject.getRowContent();

        assertThat(rowContent).contains(VALUE_1);
    }

    @Test
    void addCellContentThreeTimes_getRowContent_allCellContentIsIncludedInTheSameOrder() {
        testSubject = new SkinnyRowContentProvider();
        testSubject.addCellContent(VALUE_1);
        testSubject.addCellContent(VALUE_2);
        testSubject.addCellContent(VALUE_3);

        List<String> rowContent = testSubject.getRowContent();

        assertThat(rowContent).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void createInstanceWithVarArgs_getRowContent_constructorParametersAreIncluded() {
        testSubject = new SkinnyRowContentProvider(VALUE_1, VALUE_2, VALUE_3);

        List<String> rowContent = testSubject.getRowContent();

        assertThat(rowContent).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void createInstanceWithCollection_getRowContent_constructorParametersAreIncluded() {
        testSubject = new SkinnyRowContentProvider(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<String> rowContent = testSubject.getRowContent();

        assertThat(rowContent).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

}