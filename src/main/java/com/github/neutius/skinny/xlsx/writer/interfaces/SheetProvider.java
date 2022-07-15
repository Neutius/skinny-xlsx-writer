package com.github.neutius.skinny.xlsx.writer.interfaces;

public interface SheetProvider {

	SheetContentSupplier getSheetContentSupplier();

	String getSheetName();

}
