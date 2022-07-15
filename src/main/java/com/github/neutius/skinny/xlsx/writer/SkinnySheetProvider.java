package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetProvider;

public class SkinnySheetProvider implements SheetProvider {
	private final SheetContentSupplier sheetContentSupplier;
	private final String sheetName;

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier) {
		this(sheetContentSupplier, "");
	}

	public SkinnySheetProvider(SheetContentSupplier sheetContentSupplier, String sheetName) {
		this.sheetContentSupplier = sheetContentSupplier;
		this.sheetName = sheetName;
	}

	@Override
	public SheetContentSupplier getSheetContentSupplier() {
		return sheetContentSupplier;
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}

}
