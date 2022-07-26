package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import org.junit.jupiter.api.Test;

import java.util.List;

import static com.github.neutius.skinny.xlsx.writer.SkinnySheetProvider.EMPTY_SHEET_NAME;
import static org.assertj.core.api.Assertions.assertThat;

class SkinnySheetProviderTest {
	private static final String HEADER_1 = "Header-1";
	private static final ColumnHeaderSupplier HEADER_SUPPLIER = () -> List.of(HEADER_1);

	private static final String VALUE_1 = "value-1";
	private static final SheetContentSupplier CONTENT_SUPPLIER = () -> List.of(() -> List.of(VALUE_1));

	private static final String SHEET_NAME = "sheet name";
	private static final String NULL_NAME = null;
	private static final String BLANK_NAME = " ";

	private SkinnySheetProvider testSubject;

	@Test
	void allValuesArePresentAndValid_allValuesAreReturned() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, HEADER_SUPPLIER);

		assertThat(testSubject.getSheetContentSupplier()).isEqualTo(CONTENT_SUPPLIER);
		assertThat(testSubject.getSheetName()).isEqualTo(SHEET_NAME);
		assertThat(testSubject.getColumnHeaderSupplier()).isEqualTo(HEADER_SUPPLIER);
	}

	@Test
	void sheetContentSupplierIsNull_sheetContentSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(null, SHEET_NAME, HEADER_SUPPLIER);

		assertThat(testSubject.getSheetContentSupplier()).isNotNull();
		assertThat(testSubject.getSheetContentSupplier().get()).isNotNull().isEmpty();
	}

	@Test
	void sheetContentSupplierReturnsNull_sheetContentSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(() -> null, SHEET_NAME, HEADER_SUPPLIER);

		assertThat(testSubject.getSheetContentSupplier()).isNotNull();
		assertThat(testSubject.getSheetContentSupplier().get()).isNotNull().isEmpty();
	}

	@Test
	void noSheetNameProvided_sheetNameIsEmpty() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER);

		assertThat(testSubject.getSheetName()).isEqualTo(EMPTY_SHEET_NAME);
	}

	@Test
	void sheetNameIsNull_sheetNameIsEmpty() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, NULL_NAME, HEADER_SUPPLIER);

		assertThat(testSubject.getSheetName()).isEqualTo(EMPTY_SHEET_NAME);
		assertThat(testSubject.getSheetName()).isNotEqualTo(NULL_NAME);
	}

	@Test
	void sheetNameIsBlank_sheetNameIsEmpty() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, BLANK_NAME, HEADER_SUPPLIER);

		assertThat(testSubject.getSheetName()).isEqualTo(EMPTY_SHEET_NAME);
		assertThat(testSubject.getSheetName()).isNotEqualTo(BLANK_NAME);
	}

	@Test
	void noColumnHeaderSupplierAndNoNameProvided_columnHeaderSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER);

		assertThat(testSubject.getColumnHeaderSupplier()).isNotNull();
		assertThat(testSubject.getColumnHeaderSupplier().get()).isNotNull().isEmpty();
	}

	@Test
	void noColumnHeaderSupplierProvided_columnHeaderSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME);

		assertThat(testSubject.getColumnHeaderSupplier()).isNotNull();
		assertThat(testSubject.getColumnHeaderSupplier().get()).isNotNull().isEmpty();
	}

	@Test
	void columnHeaderSupplierIsNull_columnHeaderSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, null);

		assertThat(testSubject.getColumnHeaderSupplier()).isNotNull();
		assertThat(testSubject.getColumnHeaderSupplier().get()).isNotNull().isEmpty();
	}

	@Test
	void columnHeaderSupplierReturnsNull_columnHeaderSupplierReturnsEmptyList() {
		testSubject = new SkinnySheetProvider(CONTENT_SUPPLIER, SHEET_NAME, () -> null);

		assertThat(testSubject.getColumnHeaderSupplier()).isNotNull();
		assertThat(testSubject.getColumnHeaderSupplier().get()).isNotNull().isEmpty();
	}

}
