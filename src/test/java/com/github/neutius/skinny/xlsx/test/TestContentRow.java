package com.github.neutius.skinny.xlsx.test;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;

import java.util.ArrayList;
import java.util.List;

class TestContentRow implements ContentRowSupplier {
	private final List<String> contentRow;

	public TestContentRow() {
		contentRow = new ArrayList<>(List.of("First content cell", "    "));
		contentRow.add(null);
		contentRow.add("Last content cell");
	}

	@Override
	public List<String> get() {
		return contentRow;
	}

}
