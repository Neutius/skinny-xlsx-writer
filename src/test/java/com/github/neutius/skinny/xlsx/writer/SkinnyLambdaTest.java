package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxWorkbookProvider;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

public class SkinnyLambdaTest {

	@Test
	void useLambdasForSingleValue_valueIsPresent() {
		XlsxWorkbookProvider provider = new SkinnyWorkbookProvider(List.of(() -> List.of(() -> List.of("value-1"))));

		Workbook workbook = provider.getWorkbook();

		assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("value-1");
	}

}
