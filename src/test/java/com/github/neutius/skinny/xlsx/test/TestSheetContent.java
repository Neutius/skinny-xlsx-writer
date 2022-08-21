package com.github.neutius.skinny.xlsx.test;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.List;

public class TestSheetContent implements SheetContentSupplier {

	@Override
	public List<ContentRowSupplier> get() {
		return List.of(new TestContentRow());
	}

}
