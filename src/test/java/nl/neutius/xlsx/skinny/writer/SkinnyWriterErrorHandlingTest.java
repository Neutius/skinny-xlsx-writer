package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

class SkinnyWriterErrorHandlingTest extends AbstractSkinnyWriterTestBase {

    @Test
    void duplicateSheetNames_sheetAreAddedToWorkbook(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook(SHEET_NAME);
        writer.addSheetToWorkbook(SHEET_NAME);

        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(3);
    }

    @Test
    void duplicateSheetNames_namesWillBeGenerated(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);

        writer.addSheetToWorkbook(SHEET_NAME);
        writer.addSheetToWorkbook(SHEET_NAME);

        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isNotNull().isNotEmpty().contains(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(1).getSheetName()).isNotNull().isNotEmpty().contains(SHEET_NAME);
        assertThat(actualWorkbook.getSheetAt(2).getSheetName()).isNotNull().isNotEmpty().contains(SHEET_NAME);
    }

    @Test
    void targetFileAlreadyExists_isOverwritten(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, FILE_NAME, SHEET_NAME);
        writer.addSheetToWorkbook("sheet2");
        writer.addSheetToWorkbook("sheet3");
        writer.writeToFile();

        writer = new SkinnyWriter(targetFolder, FILE_NAME, "new sheet");

        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));

        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo("new sheet");
    }

    @Test
    void fileNameIsEmpty_nameWillBeGenerated(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, "", SHEET_NAME);

        File[] targetFolderContent = targetFolder.listFiles();
        assertThat(targetFolderContent).isNotNull().isNotEmpty().hasSize(1);

        File actualOutputFile = targetFolderContent[0];
        assertThat(actualOutputFile).exists().isFile().isNotEmpty();
        assertThat(actualOutputFile.toString()).endsWith(EXTENSION);

        String actualOutputFileName = actualOutputFile.getName();
        assertThat(actualOutputFileName).endsWith(EXTENSION).hasSizeGreaterThan(5);
        System.out.println(actualOutputFileName);

        actualWorkbook = new XSSFWorkbook(actualOutputFile);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
    }

    @Test
     void fileNameIsNull_nameWillBeGenerated(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer = new SkinnyWriter(targetFolder, null, SHEET_NAME);

        File[] targetFolderContent = targetFolder.listFiles();
        assertThat(targetFolderContent).isNotNull().isNotEmpty().hasSize(1);

        File actualOutputFile = targetFolderContent[0];
        assertThat(actualOutputFile).exists().isFile().isNotEmpty();
        assertThat(actualOutputFile.toString()).endsWith(EXTENSION);

        String actualOutputFileName = actualOutputFile.getName();
        assertThat(actualOutputFileName).endsWith(EXTENSION).hasSizeGreaterThan(5);
        System.out.println(actualOutputFileName);

        actualWorkbook = new XSSFWorkbook(actualOutputFile);
        assertThat(actualWorkbook).isNotNull().isNotEmpty().hasSize(1);
        assertThat(actualWorkbook.getSheetAt(0).getSheetName()).isEqualTo(SHEET_NAME);
    }

}