package com.github.neutius.skinny.xlsx.writer;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyColumnHeaderSupplierTest {
	private static final String HEADER_1 = "Header-1";
	private static final String HEADER_2 = "Header-2";
	private static final String HEADER_3 = "Header-3";

	private static final String EMPTY_STRING = "";

	private static final String NULL_VALUE = null;
	private static final String SPACES = "    ";
	private static final String NEW_LINES = String.format("%n%n%n%n");
	private static final String TABS = "\t\t\t\t";

	private SkinnyColumnHeaderSupplier testSubject;

	@Test
	void createInstance_rowContentIsEmpty() {
		testSubject = new SkinnyColumnHeaderSupplier();

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).isNotNull().isEmpty();
	}

	@Test
	void addColumnHeader_columnHeaderIsIncluded() {
		testSubject = new SkinnyColumnHeaderSupplier();
		testSubject.addColumnHeader(HEADER_1);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(HEADER_1);
	}

	@Test
	void addCellColumnHeaderThreeTimes_allHeadersAreIncludedInTheSameOrder() {
		testSubject = new SkinnyColumnHeaderSupplier();
		testSubject.addColumnHeader(HEADER_1);
		testSubject.addColumnHeader(HEADER_2);
		testSubject.addColumnHeader(HEADER_3);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(HEADER_1, HEADER_2, HEADER_3);
	}

	@Test
	void addNullValue_isReplacedWithEmptyString() {
		testSubject = new SkinnyColumnHeaderSupplier();
		testSubject.addColumnHeader(NULL_VALUE);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(EMPTY_STRING);
		assertThat(columnHeaders).doesNotContain(NULL_VALUE);
	}

	@Test
	void addBlankValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyColumnHeaderSupplier();
		testSubject.addColumnHeader(SPACES);
		testSubject.addColumnHeader(NEW_LINES);
		testSubject.addColumnHeader(TABS);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(columnHeaders).doesNotContain(SPACES, NEW_LINES, TABS);
	}

	@Test
	void createInstanceWithCollection_constructorParametersAreIncluded() {
		testSubject = new SkinnyColumnHeaderSupplier(Set.of(HEADER_1, HEADER_2, HEADER_3));

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactlyInAnyOrder(HEADER_1, HEADER_2, HEADER_3);
	}

	@Test
	void createInstanceWithCollectionContainingBlankAndNullValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyColumnHeaderSupplier(Arrays.asList(NULL_VALUE, SPACES, NEW_LINES, TABS));

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(columnHeaders).doesNotContain(NULL_VALUE, SPACES, NEW_LINES, TABS);
	}

	@Test
	void createInstanceWithNullCollection_isReplacedWithEmptyList() {
		testSubject = new SkinnyColumnHeaderSupplier((Set<String>) null);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).isNotNull().isEmpty();
	}

	@Test
	void createInstanceWithVarArgs_constructorParametersAreIncluded() {
		testSubject = new SkinnyColumnHeaderSupplier(HEADER_1, HEADER_2, HEADER_3);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(HEADER_1, HEADER_2, HEADER_3);
	}

	@Test
	void createInstanceWithVarArgsContainingBlankAndNullValues_areReplacedWithEmptyString() {
		testSubject = new SkinnyColumnHeaderSupplier(NULL_VALUE, SPACES, NEW_LINES, TABS);

		List<String> columnHeaders = testSubject.get();

		assertThat(columnHeaders).containsExactly(EMPTY_STRING, EMPTY_STRING, EMPTY_STRING, EMPTY_STRING);
		assertThat(columnHeaders).doesNotContain(NULL_VALUE, SPACES, NEW_LINES, TABS);
	}

}
