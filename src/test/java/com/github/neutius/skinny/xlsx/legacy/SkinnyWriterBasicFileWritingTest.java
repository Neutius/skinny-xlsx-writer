package com.github.neutius.skinny.xlsx.legacy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterBasicFileWritingTest extends AbstractSkinnyWriterTestBase {

    @Test
    void verifySetUp_targetFolderIsUsable(@TempDir File targetFolder) {
        assertThat(targetFolder).exists();
        assertThat(targetFolder).isDirectory();
        assertThat(targetFolder).isEmptyDirectory();
        assertThat(targetFolder).canRead();
        assertThat(targetFolder).canWrite();
    }

    @Test
    void createNewFile_fileExists(@TempDir File targetFolder) throws IOException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);
        assertThat(expectedFile).exists();
    }

    @Test
    void createNewFile_emptyFileIsValidXlsxFile(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

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