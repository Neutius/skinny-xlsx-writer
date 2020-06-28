package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.ss.usermodel.Sheet;

import static org.assertj.core.api.Assertions.assertThat;

abstract class AbstractSkinnyWriterTestBase {
    protected static final String EXTENSION = ".xlsx";
    protected static final String FILE_NAME = "testFile";
    protected static final String SHEET_NAME = "testSheet";

    protected SkinnyWriter writer;

    void verifyCellContent(Sheet actualSheet, int rowIndex, int columnIndex, String expectedCellContent) {
        assertThat(actualSheet.getRow(rowIndex).getCell(columnIndex).getStringCellValue()).isEqualTo(expectedCellContent);
    }
}
