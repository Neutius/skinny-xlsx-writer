package com.github.neutius.skinny.xlsx.java17.cell;

import org.apache.poi.ss.usermodel.Cell;

non-sealed class BooleanCellValue extends CellValue {

	private final boolean value;

	BooleanCellValue(boolean value) {
		this.value = value;
	}

	@Override
	public void addValue(Cell cell) {
		cell.setCellValue(value);
	}

}
