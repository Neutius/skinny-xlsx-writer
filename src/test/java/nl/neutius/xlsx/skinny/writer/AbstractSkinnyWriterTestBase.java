package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

abstract class AbstractSkinnyWriterTestBase {
    protected static final String EXTENSION = ".xlsx";
    protected static final String FILE_NAME = "testFile";
    protected static final String SHEET_NAME = "testSheet";

    protected SkinnyWriter writer;
    protected XSSFWorkbook actualWorkbook;

    void verifyCellContent(Sheet actualSheet, int rowIndex, int columnIndex, String expectedCellContent) {
        assertThat(actualSheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue()).isEqualTo(expectedCellContent);
    }

    @AfterEach
    void closeActualWorkbook() throws IOException {
        if (actualWorkbook != null) {
            actualWorkbook.close();
        }
    }
}
