package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Optional;

/**
 * This class writes the content of an in memory Apache POI Workbook to an .xlsx file on the file system.
 * <p>
 * After any write action, feedback on the success of that write action is available in the form of a boolean.
 * Any IOException that was caught during that write action is also available, and might be rethrown by the client if so desired.
 */
public class SkinnyFileWriter {
    private static final Logger LOG = LoggerFactory.getLogger(SkinnyFileWriter.class);

    private boolean lastWriteSuccessful;
    private Exception lastWriteException;

    /**
     * Provides feedback on the success of the previous write attempt.
     * <p>
     * Returns true if no Exception occurred during the previous write action.
     * Returns false if an Exception did occur, or if no write action has been performed.
     *
     * @return A best-effort indication of whether your content has been written to the file system.
     */
    public boolean isLastWriteSuccessful() {
        return lastWriteSuccessful;
    }

    /**
     * Any Exception that occurs is stored and available for later retrieval, until another write action is performed.
     *
     * @return the Exception that may or may not have occurred during the previous write attempt.
     */
    public Optional<Exception> getLastWriteException() {
        return Optional.ofNullable(lastWriteException);
    }

    /**
     * Write the content of a Workbook to a File on the file system.
     * Any exception that occurs wil be caught, logged and available for retrieval until this method is called again.
     *
     * @param content    The content to be written to the file system.
     * @param outputFile The location of the file to be written.
     *                   If this file already exists, a new file within the same directory will be created.
     */
    public void write(Workbook content, File outputFile) {
        try {
            writeContentToFileSystem(content, sanitizeOutputFile(outputFile));
            lastWriteSuccessful = true;
            lastWriteException = null;
        } catch (Exception e) {
            LOG.warn("An exception occurred while writing to the file system", e);
            lastWriteSuccessful = false;
            lastWriteException = e;
        }
    }

    private void writeContentToFileSystem(Workbook content, File outputFile) throws IOException {
        outputFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(outputFile);
        content.write(outputStream);
        outputStream.close();
    }

    private static File sanitizeOutputFile(File outputFile) {
        if (outputFile.exists()) {
            File actualOutputFile = getActualOutputFile(outputFile);
            LOG.info("File {} already exists. Workbook content will be written to new file {}", outputFile, actualOutputFile);
            return actualOutputFile;
        }
        return outputFile;
    }

    private static File getActualOutputFile(File outputFile) {
        return new File(outputFile.getParentFile(), getFileName());
    }

    private static String getFileName() {
        return "output-at-" + getTimeStamp() + ".xlsx";
    }

    private static String getTimeStamp() {
        return new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss").format(new Date());
    }

}
