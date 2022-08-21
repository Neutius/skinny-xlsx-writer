package com.github.neutius.skinny.xlsx.test;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;

import java.util.ArrayList;
import java.util.List;

public class TestColumnHeaders implements ColumnHeaderSupplier {
	private final List<String> columnHeaders;

	public TestColumnHeaders() {
		columnHeaders = new ArrayList<>(List.of("First column", "    "));
		columnHeaders.add(null);
		columnHeaders.add("Last column");
	}

	@Override
	public List<String> get() {
		return columnHeaders;
	}

}
