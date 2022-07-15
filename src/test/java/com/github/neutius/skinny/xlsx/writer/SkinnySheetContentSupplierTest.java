package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnySheetContentSupplierTest {
    private static final String VALUE_1 = "value1";
    private static final String VALUE_2 = "value2";
    private static final String VALUE_3 = "value3";
    private static final String VALUE_4 = "value4";
    private static final String VALUE_5 = "value5";
    private static final String VALUE_6 = "value6";

    private SkinnySheetContentSupplier testSubject;

    @Test
    void createInstance_resultIsEmpty() {
        testSubject = new SkinnySheetContentSupplier();

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).isNotNull().isEmpty();
    }

    @Test
    void addRowContentSupplier_inputIsIncluded() {
        testSubject = new SkinnySheetContentSupplier();
        SkinnyRowContentSupplier rowContentSupplier = new SkinnyRowContentSupplier();
        testSubject.addRowContentSupplier(rowContentSupplier);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).contains(rowContentSupplier);
    }

    @Test
    void addRowContentWithSupplier_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addRowContentSupplier(new SkinnyRowContentSupplier(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers.get(0).get()).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(VALUE_1, VALUE_2, VALUE_3);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void createInstanceWithRowContentAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(List.of(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6)));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(rowContentSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

    @Test
    void createInstanceWithRowContentAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(rowContentSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

}
