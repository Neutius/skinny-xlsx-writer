package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class SkinnyFileWriterTest {
    private final IOException testException = new IOException("Test Exception");
    private SkinnyFileWriter testSubject;

    @BeforeEach
    void setup() {
        testSubject = new SkinnyFileWriter();
    }

    @Test
    void noWriteActionPerformed_getSuccess_notSuccessful() {
        assertThat(testSubject.isLastWriteSuccessful()).isFalse();
        assertThat(testSubject.getLastWriteException()).isEmpty();
    }

    @Test
    void unsuccessfulWriteActionPerformed_getSuccess_notSuccessful() throws IOException {
        Workbook content = mock(Workbook.class);
        File outputFile = mock(File.class);
        when(outputFile.getParentFile()).thenReturn(mock(File.class));
        when(outputFile.createNewFile()).thenThrow(testException);

        testSubject.write(content, outputFile);

        assertThat(testSubject.isLastWriteSuccessful()).isFalse();
        assertThat(testSubject.getLastWriteException()).isNotEmpty().contains(testException);
    }

    @Test
    void successfulWriteActionPerformed_getSuccess_isSuccessful(@TempDir File targetFolder) {
        Workbook content = mock(Workbook.class);
        File outputFile = new File(targetFolder, "test.xlsx");

        testSubject.write(content, outputFile);

        assertThat(testSubject.isLastWriteSuccessful()).isTrue();
        assertThat(testSubject.getLastWriteException()).isEmpty();
    }

    @Test
    void fileDoesNotExist_writeToFile_fileIsCreated(@TempDir File targetFolder) {
        Workbook content = mock(Workbook.class);
        File outputFile = new File(targetFolder, "test.xlsx");

        assertThat(outputFile).doesNotExist();
        assertThat(targetFolder.listFiles()).hasSize(0);

        testSubject.write(content, outputFile);

        assertThat(outputFile).exists();
        assertThat(targetFolder.listFiles()).hasSize(1);
    }

    @Test
    void fileAlreadyExists_writeToFile_anotherFileIsCreated(@TempDir File targetFolder) throws IOException {
        Workbook content = mock(Workbook.class);
        File outputFile = new File(targetFolder, "test.xlsx");
        outputFile.createNewFile();

        assertThat(outputFile).exists();
        assertThat(targetFolder.listFiles()).hasSize(1);

        testSubject.write(content, outputFile);

        assertThat(outputFile).exists();
        assertThat(targetFolder.listFiles()).hasSize(2);
    }

    @Test
    void outputFileInNonExistentDirectory_writeToFile_bothDirectoryAndFileAreCreated(@TempDir File targetFolder) {
        Workbook content = mock(Workbook.class);
        File nonExistentDirectory = new File(targetFolder, "non-existent");
        File outputFile = new File(nonExistentDirectory, "test.xlsx");

        assertThat(targetFolder.listFiles()).hasSize(0);
        assertThat(nonExistentDirectory).doesNotExist();
        assertThat(outputFile).doesNotExist();

        testSubject.write(content, outputFile);

        assertThat(targetFolder.listFiles()).hasSize(1);
        assertThat(nonExistentDirectory).exists();
        assertThat(outputFile).exists();
    }

}
