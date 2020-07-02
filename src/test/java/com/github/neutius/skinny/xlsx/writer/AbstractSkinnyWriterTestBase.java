package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;

abstract class AbstractSkinnyWriterTestBase {
    protected static final String EXTENSION = ".xlsx";
    protected static final String FILE_NAME = "testFile";
    protected static final String SHEET_NAME = "testSheet";

    protected SkinnyWriter writer;
    protected XSSFWorkbook actualWorkbook;

    @AfterEach
    void closeActualWorkbook() throws IOException {
        if (actualWorkbook != null) {
            actualWorkbook.close();
        }
    }

    protected void writeAndReadActualWorkbook(@TempDir File targetFolder) throws IOException, InvalidFormatException {
        writer.writeToFile();
        actualWorkbook = new XSSFWorkbook(new File(targetFolder, FILE_NAME + EXTENSION));
    }
}
