package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;

import java.util.Collections;

public class SkinnySheetProvider implements SheetProvider {
	public static final String EMPTY_SHEET_NAME = "";
	public static final ColumnHeaderSupplier EMPTY_COLUMN_HEADER = Collections::emptyList;
	public static final SheetContentSupplier EMPTY_SHEET_CONTENT = Collections::emptyList;
	public static final SheetProvider EMPTY_SHEET = new SkinnySheetProvider(EMPTY_SHEET_CONTENT, EMPTY_SHEET_NAME, EMPTY_COLUMN_HEADER);

	private final String sheetName;
	private final ColumnHeaderSupplier columnHeaderSupplier;
	private final SheetContentSupplier sheetContentSupplier;

	public static SheetProvider empty() {
		return EMPTY_SHEET;
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}

	@Override
	public ColumnHeaderSupplier getColumnHeaderSupplier() {
		return columnHeaderSupplier;
	}

	@Override
	public SheetContentSupplier getSheetContentSupplier() {
		return sheetContentSupplier;
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier) {
		this(sheetContentSupplier, EMPTY_SHEET_NAME);
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier, String sheetName) {
		this(sheetContentSupplier, sheetName, EMPTY_COLUMN_HEADER);
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier, String sheetName, ColumnHeaderSupplier headerSupplier) {
		this.sheetContentSupplier = sanitizeContentSupplier(sheetContentSupplier);
		this.sheetName = sanitizeSheetName(sheetName);
		this.columnHeaderSupplier = sanitizeHeaderSupplier(headerSupplier);
	}

	private static SheetContentSupplier sanitizeContentSupplier(SheetContentSupplier contentSupplier) {
		return contentSupplier == null || contentSupplier.get() == null ? EMPTY_SHEET_CONTENT : contentSupplier;
	}

	private static String sanitizeSheetName(String sheetName) {
		return sheetName == null || sheetName.isBlank() ? EMPTY_SHEET_NAME : sheetName;
	}

	private static ColumnHeaderSupplier sanitizeHeaderSupplier(ColumnHeaderSupplier headerSupplier) {
		return headerSupplier == null || headerSupplier.get() == null ? EMPTY_COLUMN_HEADER : headerSupplier;
	}

}
