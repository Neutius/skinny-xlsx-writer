package com.github.neutius.skinny.xlsx.test;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;

public class TestSheet implements SheetProvider {
	private final String sheetName;
	private final ColumnHeaderSupplier columnHeaderSupplier;
	private final SheetContentSupplier sheetContentSupplier;

	public TestSheet(String sheetName) {
		this(sheetName, null, null);
	}

	public TestSheet(SheetContentSupplier sheetContentSupplier) {
		this(null, null, sheetContentSupplier);
	}

	public TestSheet(String sheetName, ColumnHeaderSupplier columnHeaderSupplier, SheetContentSupplier sheetContentSupplier) {
		this.sheetName = sheetName;
		this.columnHeaderSupplier = columnHeaderSupplier;
		this.sheetContentSupplier = sheetContentSupplier;
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

}
