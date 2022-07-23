package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
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

        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers).contains(rowContentSupplier);
    }

    @Test
    void addRowContentWithSupplier_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addRowContentSupplier(new SkinnyRowContentSupplier(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addNullValueForRowContentSupplier_isReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addRowContentSupplier(null);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).doesNotContain((RowContentSupplier) null);
        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).isNotNull().isEmpty();
    }

    @Test
    void addRowContentAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsNullCollection_isReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow((Set<String>) null);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).doesNotContain((RowContentSupplier) null);
        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).isNotNull().isEmpty();
    }

    @Test
    void addRowContentAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(VALUE_1, VALUE_2, VALUE_3);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addRowContentAsNullVarArgs_resultContainsEmptyStrings() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(null, null, null);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(1);
        assertThat(rowContentSuppliers.get(0).get()).doesNotContain(VALUE_1, VALUE_2, VALUE_3);
        assertThat(rowContentSuppliers.get(0).get()).containsExactly("", "", "");
    }

    @Test
    void createInstanceWithRowContentAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(List.of(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6)));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(2);
        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(rowContentSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

    @Test
    void createInstanceWithRowContentAsCollectionContainingNull_nullReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier(Arrays.asList(
                () -> null,
                null));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(2);
        assertThat(rowContentSuppliers.get(0).get()).isNotNull().isEmpty();
        assertThat(rowContentSuppliers.get(1).get()).isNotNull().isEmpty();
    }

    @Test
    void createInstanceWithRowContentAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6));

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(2);
        assertThat(rowContentSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(rowContentSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

    @Test
    void createInstanceWithRowContentAsVarArgsContainingNull_nullReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier(
                () -> null,
                null);

        List<RowContentSupplier> rowContentSuppliers = testSubject.get();

        assertThat(rowContentSuppliers).hasSize(2);
        assertThat(rowContentSuppliers.get(0).get()).isNotNull().isEmpty();
        assertThat(rowContentSuppliers.get(1).get()).isNotNull().isEmpty();
    }

}
