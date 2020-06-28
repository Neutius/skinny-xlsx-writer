package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;

import java.io.File;
import java.io.IOException;

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

    @EnabledIfSystemProperty(matches = "true", named = "test.local")
    @Test
    void verifySetUp_targetFolderIsUsable() {
        assertThat(targetFolder).exists();
        assertThat(targetFolder).isDirectory();
        assertThat(targetFolder).isEmptyDirectory();
        assertThat(targetFolder).canRead();
        assertThat(targetFolder).canWrite();
    }

    @EnabledIfSystemProperty(matches = "true", named = "test.local")
    @Test
    void createNewFile_fileExists() throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        assertThat(expectedFile).exists();
    }

    @EnabledIfSystemProperty(matches = "true", named = "test.local")
    @Test
    void createNewFile_emptyFileIsValidXlsxFile() throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.createNewXlsxFile();

        File targetFile = new File(targetFolder, FILE_NAME + EXTENSION);
        actualWorkbook = new XSSFWorkbook(targetFile);
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

}