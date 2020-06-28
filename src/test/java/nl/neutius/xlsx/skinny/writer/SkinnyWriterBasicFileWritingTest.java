package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterBasicFileWritingTest extends AbstractSkinnyWriterTestBase {

    protected File targetFolder;
    
    @BeforeEach
    void setUpTargetFolder() {
        targetFolder = new File("D:\\dev\\test\\excel");
        File[] fileArray = targetFolder.listFiles();

        if (fileArray != null && fileArray.length > 0) {
            for (File file : fileArray) {
                assertThat(file.delete()).isTrue();
            }
        }
    }

    @Test
    void verifySetUp_targetFolderIsUsable() {
        assertThat(targetFolder).exists();
        assertThat(targetFolder).isDirectory();
        assertThat(targetFolder).isEmptyDirectory();
        assertThat(targetFolder).canRead();
        assertThat(targetFolder).canWrite();
    }

    @Test
    void createNewFile_fileExists() throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        assertThat(expectedFile).exists();
    }

    @Test
    void createNewFile_emptyFileIsValidXlsxFile() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File targetFile = new File(targetFolder, FILE_NAME + EXTENSION);
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(targetFile);
        XSSFSheet actualSheet = actualWorkbook.getSheetAt(0);

        assertThat(actualSheet).isNotNull();
        assertThat(actualSheet).isEmpty();
        assertThat(actualSheet).hasSize(0);
        assertThat(actualSheet.getActiveCell()).isNull();
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(-1);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(-1);
        assertThat(actualSheet.getCellComments()).isEmpty();
        assertThat(actualSheet.getSheetName()).isEqualTo(SHEET_NAME);
        assertThat(actualSheet).isEqualTo(actualWorkbook.getSheet(SHEET_NAME));
    }

    @Test
    void addContent_fileHasContent() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();
        writer.addRowToCurrentSheet(List.of("entry"));
        writer.writeToFile();
        XSSFWorkbook actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        Sheet actualSheet = actualWorkbook.getSheet(SHEET_NAME);
        assertThat(actualSheet).isNotEmpty();
        assertThat(actualSheet).hasSize(1);
        assertThat(actualSheet.getFirstRowNum()).isEqualTo(0);
        assertThat(actualSheet.getLastRowNum()).isEqualTo(0);

        Row actualRow = actualSheet.getRow(0);
        assertThat(actualRow).isNotNull();
        assertThat(actualRow).isNotEmpty();
        assertThat(actualRow).hasSize(1);
        assertThat((int) actualRow.getFirstCellNum()).isEqualTo(0);
        assertThat((int) actualRow.getLastCellNum()).isEqualTo(1);
        assertThat(actualRow.getPhysicalNumberOfCells()).isEqualTo(1);
        assertThat(actualRow.getRowNum()).isEqualTo(0);

        Cell actualCell = actualRow.getCell(0);
        assertThat(actualCell).isNotNull();
        assertThat(actualCell.getSheet()).isEqualTo(actualSheet);
        assertThat(actualCell.getStringCellValue()).isEqualTo("entry");
        assertThat(actualCell.getColumnIndex()).isEqualTo(0);
        assertThat(actualCell.getRowIndex()).isEqualTo(0);
    }

}