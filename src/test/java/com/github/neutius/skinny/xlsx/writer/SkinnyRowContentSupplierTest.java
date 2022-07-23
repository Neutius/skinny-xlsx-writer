package com.github.neutius.skinny.xlsx.writer;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyRowContentSupplierTest {
    private static final String VALUE_1 = "value1";
    private static final String VALUE_2 = "value2";
    private static final String VALUE_3 = "value3";

    private static final String EMPTY_STRING = "";
    private static final String SPACES = "    ";
    private static final String NEW_LINES = String.format("%n%n%n%n");
    private static final String TABS = "\t\t\t\t";

    private SkinnyRowContentSupplier testSubject;

    @Test
    void createInstance_rowContentIsEmpty() {
        testSubject = new SkinnyRowContentSupplier();

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).isNotNull().isEmpty();
    }

    @Test
    void addCellContent_cellContentIsIncluded() {
        testSubject = new SkinnyRowContentSupplier();
        testSubject.addCellContent(VALUE_1);

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).contains(VALUE_1);
    }

    @Test
    void addCellContentThreeTimes_allCellContentIsIncludedInTheSameOrder() {
        testSubject = new SkinnyRowContentSupplier();
        testSubject.addCellContent(VALUE_1);
        testSubject.addCellContent(VALUE_2);
        testSubject.addCellContent(VALUE_3);

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Test
    void addNullValue_isReplacedWithEmptyString() {
        testSubject = new SkinnyRowContentSupplier();
        testSubject.addCellContent(null);

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).containsExactly(EMPTY_STRING);
        assertThat(rowContent).doesNotContain((String) null);
    }

    @Test
    void addBlankValues_areReplacedWithEmptyString() {
        testSubject = new SkinnyRowContentSupplier();
        testSubject.addCellContent(SPACES);
        testSubject.addCellContent(NEW_LINES);
        testSubject.addCellContent(TABS);

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
        assertThat(rowContent).doesNotContain(SPACES, NEW_LINES, TABS);
    }

    @Test
    void createInstanceWithVarArgs_constructorParametersAreIncluded() {
        testSubject = new SkinnyRowContentSupplier(VALUE_1, VALUE_2, VALUE_3);

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).containsExactly(VALUE_1, VALUE_2, VALUE_3);
    }

    @Disabled("TODO write test -> adjust implementation if needed - GvdNL 15-07-2022")
    @Test
    void createInstanceWithVarArgsContainingBlankAndNullValues_areReplacedWithEmptyString() {
        // TODO write test -> adjust implementation if needed - GvdNL 15-07-2022
    }

    @Test
    void createInstanceWithCollection_constructorParametersAreIncluded() {
        testSubject = new SkinnyRowContentSupplier(Set.of(VALUE_1, VALUE_2, VALUE_3));

        List<String> rowContent = testSubject.get();

        assertThat(rowContent).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
    }

    @Disabled("TODO write test -> adjust implementation if needed - GvdNL 15-07-2022")
    @Test
    void createInstanceWithCollectionContainingBlankAndNullValues_areReplacedWithEmptyString() {
        // TODO write test -> adjust implementation if needed - GvdNL 15-07-2022
    }

}
