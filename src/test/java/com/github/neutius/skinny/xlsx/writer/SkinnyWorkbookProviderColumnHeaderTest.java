package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWorkbookProviderColumnHeaderTest {
	private static final String HEADER_1 = "Header-1";
	private static final String VALUE_1 = "value-1";
	private static final String SHEET_NAME = "sheet name";

	private SkinnyWorkbookProvider testSubject;

	@Test
	void addColumnHeader_isPresent() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetContentSupplier contentSupplier = () -> List.of(Collections::emptyList);
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(HEADER_1);
	}

	@Test
	void addColumnHeaderAndContentRow_bothArePresent() {
		ColumnHeaderSupplier headerSupplier = () -> List.of(HEADER_1);
		SheetContentSupplier contentSupplier = () -> List.of(() -> List.of(VALUE_1));
		SheetProvider sheetProvider = new SkinnySheetProvider(contentSupplier, SHEET_NAME, headerSupplier);
		testSubject = new SkinnyWorkbookProvider(List.of(sheetProvider));

		Workbook workbook = testSubject.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo(HEADER_1);
		assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo(VALUE_1);
	}



	/*
	TODO add tests and functionality - GvdNL 26-07-2022

	- handling of null values

	- addColumnHeader_fontIsBold
	- addColumnHeader_freezePaneIsApplied
	- addColumnHeader_filterIsApplied

	- addColumnHeaderAndContentRow_contentRowFontIsNotBold
	- addColumnHeaderAndContentRow_contentRowHasNoFreezePane
	- addColumnHeaderAndContentRow_contentRowHasNoFilter

	- configuration options for bold font, freeze pane and filter?

	- columnHeadersHaveDifferentWidth_sheetColumnsHaveDifferentWidth

	Maybe here or somewhere else?
	- cellValuesHaveDifferentWidthOnlyAfter99ContentRows_sheetColumnsHaveSameWidth
	- cellValuesHaveDifferentWidthOnlyAfter98ContentRows_sheetColumnsHaveDifferentWidth

	 */


}
