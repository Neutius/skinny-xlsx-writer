package com.github.neutius.skinny.xlsx.legacy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SkinnyWriterConstructorTest extends AbstractSkinnyWriterTestBase {


    @Test
    void useConstructorWithoutFirstSheetName_noFileIsCreated(@TempDir File targetFolder) {
        assertThat(targetFolder).isEmptyDirectory();

        writer = new SkinnyWriter(targetFolder, FILE_NAME);

        assertThat(targetFolder).isEmptyDirectory();
    }

    @Test
    void noFirstSheet_writeFile_fileHasNoSheets(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        File expectedFile = new File(targetFolder, FILE_NAME + EXTENSION);

        writer = new SkinnyWriter(targetFolder, FILE_NAME);
        writer.writeToFile();

        assertThat(targetFolder).isNotEmptyDirectory();
        assertThat(expectedFile).isNotNull().exists().isFile().isNotEmpty();

        actualWorkbook = new XSSFWorkbook(expectedFile);
        assertThat(actualWorkbook).hasSize(0);
    }

    @Test
    void noFirstSheet_addColumnHeaderRow_exceptionIsThrown(@TempDir File targetFolder) {
        writer = new SkinnyWriter(targetFolder, FILE_NAME);

        assertThatThrownBy(() -> writer.addColumnHeaderRowToCurrentSheet(List.of("header1", "header2")))
                .isInstanceOf(NullPointerException.class);
    }

    @Test
    void noFirstSheet_addContentRow_exceptionIsThrown(@TempDir File targetFolder) {
        writer = new SkinnyWriter(targetFolder, FILE_NAME);

        assertThatThrownBy(() -> writer.addRowToCurrentSheet(List.of("content1", "content2")))
                .isInstanceOf(NullPointerException.class);


    }

}