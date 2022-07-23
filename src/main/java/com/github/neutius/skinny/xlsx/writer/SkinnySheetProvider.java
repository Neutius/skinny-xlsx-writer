package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;

public class SkinnySheetProvider implements SheetProvider {
	// Should these fields be final?
	// Perhaps adding a zero-argument constructor and setters would add flexibility? - GvdNL 23-07-2022
	private final SheetContentSupplier sheetContentSupplier;
	private final String sheetName;

	@Override
	public SheetContentSupplier getSheetContentSupplier() {
		return sheetContentSupplier;
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier) {
		this(sheetContentSupplier, "");
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier, String sheetName) {
		this.sheetContentSupplier = sheetContentSupplier;
		this.sheetName = sheetName;
	}

}