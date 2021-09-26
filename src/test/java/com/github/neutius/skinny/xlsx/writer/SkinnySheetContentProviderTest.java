package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxRowContentProvider;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnySheetContentProviderTest {
    private static final String VALUE_1 = "value1";
    private static final String VALUE_2 = "value2";
    private static final String VALUE_3 = "value3";

    private SkinnySheetContentProvider testSubject;

    @Test
    void createInstance_getRowContentProviders_resultIsEmpty() {
        testSubject = new SkinnySheetContentProvider();

        List<XlsxRowContentProvider> rowContentProviders = testSubject.getRowContentProviders();

        assertThat(rowContentProviders).isNotNull().isEmpty();
    }

    @Test
    void addRowContentProvider_getRowContentProviders_inputIsIncluded() {
        testSubject = new SkinnySheetContentProvider();
        SkinnyRowContentProvider rowContentProvider = new SkinnyRowContentProvider();
        testSubject.addRowContentProvider(rowContentProvider);

        List<XlsxRowContentProvider> rowContentProviders = testSubject.getRowContentProviders();

        assertThat(rowContentProviders).contains(rowContentProvider);
    }

    @Test
    void addRowContentWithProvider_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentProvider();
        testSubject.addRowContentProvider(new SkinnyRowContentProvider(VALUE_1, VALUE_2, VALUE_3));

        List<XlsxRowContentProvider> rowContentProviders = testSubject.getRowContentProviders();

        assertThat(rowContentProviders.get(0).getRowContent()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsVarArgs_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentProvider();
        testSubject.addContentRow(VALUE_1, VALUE_2, VALUE_3);

        List<XlsxRowContentProvider> rowContentProviders = testSubject.getRowContentProviders();

        assertThat(rowContentProviders.get(0).getRowContent()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsCollection_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentProvider();
        testSubject.addContentRow(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<XlsxRowContentProvider> rowContentProviders = testSubject.getRowContentProviders();

        assertThat(rowContentProviders.get(0).getRowContent()).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

}