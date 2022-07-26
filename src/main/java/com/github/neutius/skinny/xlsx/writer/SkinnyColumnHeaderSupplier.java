package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ColumnHeaderSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnyColumnHeaderSupplier implements ColumnHeaderSupplier {
	private final List<String> columnHeaders = new ArrayList<>();

	@Override
	public List<String> get() {
		return columnHeaders;
	}

	public SkinnyColumnHeaderSupplier() {
	}

	public SkinnyColumnHeaderSupplier(Collection<String> initialContent) {
		if (initialContent != null) {
			initialContent.forEach(this::addHeaderToColumnHeaderRow);
		}
	}

	public SkinnyColumnHeaderSupplier(String... initialContent) {
		Arrays.asList(initialContent).forEach(this::addHeaderToColumnHeaderRow);
	}

	public void addColumnHeader(String header) {
		addHeaderToColumnHeaderRow(header);
	}

	private void addHeaderToColumnHeaderRow(String header) {
		columnHeaders.add(sanitizeHeader(header));
	}

	private static String sanitizeHeader(String header) {
		return header == null || header.isBlank() ? "" : header;
	}

}
