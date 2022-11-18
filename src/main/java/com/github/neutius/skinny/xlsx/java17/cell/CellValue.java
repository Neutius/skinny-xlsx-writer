package com.github.neutius.skinny.xlsx.java17.cell;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

/**
 * This should become a sealed class, with a limited and known set of implementations: one of each data-type supported by
 * Excel and Apache POI.
 *
 *
 */
public abstract sealed class CellValue permits BooleanCellValue {

	public static CellValue of(boolean value) {
		return new BooleanCellValue(value);
	}

	public abstract void addValue(Cell cell);

	static void discoverSupportedTypes(Cell cell, CellValue value) {
		boolean booleanCellValue = true;
		cell.setCellValue(booleanCellValue);
		cell.setCellValue(12.3);
		cell.setCellValue("Normal text");

		cell.setCellValue(new XSSFRichTextString());

		LocalDate localDate = LocalDate.now();
		cell.setCellValue(localDate);

		LocalDateTime localDateTime = LocalDateTime.now();
		cell.setCellValue(localDateTime);

		Date date = Date.from(Instant.now());
		cell.setCellValue(date);

		Calendar calendar = Calendar.getInstance();
		cell.setCellValue(calendar);

	}

}
