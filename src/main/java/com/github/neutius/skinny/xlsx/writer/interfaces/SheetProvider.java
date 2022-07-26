package com.github.neutius.skinny.xlsx.writer.interfaces;

public interface SheetProvider {

	String getSheetName();

	ColumnHeaderSupplier getColumnHeaderSupplier();

	SheetContentSupplier getSheetContentSupplier();

}
