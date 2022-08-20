package com.github.neutius.skinny.xlsx.writer;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyContentRowSupplierTest {
	private static final String VALUE_1 = "value1";
	private static final String VALUE_2 = "value2";
	private static final String VALUE_3 = "value3";

	private static final String EMPTY_STRING = "";

	private static final String NULL_VALUE = null;
	private static final String SPACES = "    ";
	private static final String NEW_LINES = String.format("%n%n%n%n");
	private static final String TABS = "\t\t\t\t";

	private SkinnyContentRowSupplier testSubject;

	@Test
	void createInstance_rowIsEmpty() {
		testSubject = new SkinnyContentRowSupplier();

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).isNotNull().isEmpty();
	}

	@Test
	void addCellContent_cellContentIsIncluded() {
		testSubject = new SkinnyContentRowSupplier();
		testSubject.addCellContent(VALUE_1);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(VALUE_1);
	}

	@Test
	void addCellContentThreeTimes_allCellContentIsIncludedInTheSameOrder() {
		testSubject = new SkinnyContentRowSupplier();
		testSubject.addCellContent(VALUE_1);
		testSubject.addCellContent(VALUE_2);
		testSubject.addCellContent(VALUE_3);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(VALUE_1, VALUE_2, VALUE_3);
	}

	@Test
	void addNullValue_isReplacedWithEmptyString() {
		testSubject = new SkinnyContentRowSupplier();
		testSubject.addCellContent(NULL_VALUE);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(EMPTY_STRING);
		assertThat(contentRow).doesNotContain(NULL_VALUE);
	}

	@Test
	void addBlankValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyContentRowSupplier();
		testSubject.addCellContent(SPACES);
		testSubject.addCellContent(NEW_LINES);
		testSubject.addCellContent(TABS);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(contentRow).doesNotContain(SPACES, NEW_LINES, TABS);
	}

	@Test
	void createInstanceWithCollection_constructorParametersAreIncluded() {
		testSubject = new SkinnyContentRowSupplier(Set.of(VALUE_1, VALUE_2, VALUE_3));

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactlyInAnyOrder(VALUE_1, VALUE_2, VALUE_3);
	}

	@Test
	void createInstanceWithCollectionContainingBlankAndNullValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyContentRowSupplier(Arrays.asList(NULL_VALUE, SPACES, NEW_LINES, TABS));

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(contentRow).doesNotContain(NULL_VALUE, SPACES, NEW_LINES, TABS);
	}

	@Test
	void createInstanceWithNullCollection_isReplacedWithEmptyList() {
		testSubject = new SkinnyContentRowSupplier((Set<String>) null);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).isNotNull().isEmpty();
	}

	@Test
	void createInstanceWithVarArgs_constructorParametersAreIncluded() {
		testSubject = new SkinnyContentRowSupplier(VALUE_1, VALUE_2, VALUE_3);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(VALUE_1, VALUE_2, VALUE_3);
	}

	@Test
	void createInstanceWithVarArgsContainingBlankAndNullValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyContentRowSupplier(NULL_VALUE, SPACES, NEW_LINES, TABS);

		List<String> contentRow = testSubject.get();

		assertThat(contentRow).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(contentRow).doesNotContain(NULL_VALUE, SPACES, NEW_LINES, TABS);
	}

}
