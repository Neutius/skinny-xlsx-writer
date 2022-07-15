package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnySheetContentProviderTest {
    private static final String VALUE_1 = "value1";
    private static final String VALUE_2 = "value2";
    private static final String VALUE_3 = "value3";

    private SkinnySheetContentSupplier testSubject;

    @Test
    void createInstance_getRowContentProviders_resultIsEmpty() {
        testSubject = new SkinnySheetContentSupplier();

        List<RowContentSupplier> rowContentProviders = testSubject.get();

        assertThat(rowContentProviders).isNotNull().isEmpty();
    }

    @Test
    void addRowContentProvider_getRowContentProviders_inputIsIncluded() {
        testSubject = new SkinnySheetContentSupplier();
        SkinnyRowContentSupplier rowContentProvider = new SkinnyRowContentSupplier();
        testSubject.addRowContentProvider(rowContentProvider);

        List<RowContentSupplier> rowContentProviders = testSubject.get();

        assertThat(rowContentProviders).contains(rowContentProvider);
    }

    @Test
    void addRowContentWithProvider_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addRowContentProvider(new SkinnyRowContentSupplier(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentProviders = testSubject.get();

        assertThat(rowContentProviders.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsVarArgs_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(VALUE_1, VALUE_2, VALUE_3);

        List<RowContentSupplier> rowContentProviders = testSubject.get();

        assertThat(rowContentProviders.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsCollection_getRowContentProviders_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentProviders = testSubject.get();

        assertThat(rowContentProviders.get(0).get()).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

}