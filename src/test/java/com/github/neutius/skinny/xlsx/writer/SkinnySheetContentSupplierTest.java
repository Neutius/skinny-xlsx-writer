package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
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

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).isNotNull().isEmpty();
    }

    @Test
    void addContentRowSupplier_inputIsIncluded() {
        testSubject = new SkinnySheetContentSupplier();
        SkinnyContentRowSupplier contentRowSupplier = new SkinnyContentRowSupplier();
        testSubject.addContentRowSupplier(contentRowSupplier);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers).contains(contentRowSupplier);
    }

    @Test
    void addContentRowWithSupplier_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRowSupplier(new SkinnyContentRowSupplier(VALUE_1, VALUE_2, VALUE_3));

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addNullValueForContentRowSupplier_isReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRowSupplier(null);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).doesNotContain((ContentRowSupplier) null);
        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).isNotNull().isEmpty();
    }

    @Test
    void addContentRowAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addContentRowAsNullCollection_isReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow((Set<String>) null);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).doesNotContain((ContentRowSupplier) null);
        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).isNotNull().isEmpty();
    }

    @Test
    void addContentRowAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(VALUE_1, VALUE_2, VALUE_3);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addContentRowAsNullVarArgs_resultContainsEmptyStrings() {
        testSubject = new SkinnySheetContentSupplier();
        testSubject.addContentRow(null, null, null);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(1);
        assertThat(contentRowSuppliers.get(0).get()).doesNotContain(VALUE_1, VALUE_2, VALUE_3);
        assertThat(contentRowSuppliers.get(0).get()).containsExactly("", "", "");
    }

    @Test
    void createInstanceWithContentRowAsCollection_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(List.of(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6)));

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(2);
        assertThat(contentRowSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(contentRowSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

    @Test
    void createInstanceWithContentRowAsCollectionContainingNull_nullReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier(Arrays.asList(
                () -> null,
                null));

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(2);
        assertThat(contentRowSuppliers.get(0).get()).isNotNull().isEmpty();
        assertThat(contentRowSuppliers.get(1).get()).isNotNull().isEmpty();
    }

    @Test
    void createInstanceWithContentRowAsVarArgs_sameValuesAreReturned() {
        testSubject = new SkinnySheetContentSupplier(
                () -> List.of(VALUE_1, VALUE_2, VALUE_3),
                () -> List.of(VALUE_4, VALUE_5, VALUE_6));

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(2);
        assertThat(contentRowSuppliers.get(0).get()).containsExactly(VALUE_1, VALUE_2, VALUE_3);
        assertThat(contentRowSuppliers.get(1).get()).containsExactly(VALUE_4, VALUE_5, VALUE_6);
    }

    @Test
    void createInstanceWithContentRowAsVarArgsContainingNull_nullReplacedWithEmptyList() {
        testSubject = new SkinnySheetContentSupplier(
                () -> null,
                null);

        List<ContentRowSupplier> contentRowSuppliers = testSubject.get();

        assertThat(contentRowSuppliers).hasSize(2);
        assertThat(contentRowSuppliers.get(0).get()).isNotNull().isEmpty();
        assertThat(contentRowSuppliers.get(1).get()).isNotNull().isEmpty();
    }

}
