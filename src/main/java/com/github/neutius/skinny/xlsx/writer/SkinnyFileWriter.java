package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxFileWriter;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Optional;

/**
 * This class writes the content of an in memory Apache POI Workbook to an .xlsx file on the file system.
 * <p>
 * After any write action, feedback on the success of that write action is available in the form of a boolean.
 * Any Exception that was caught during that write action is also available, and can be rethrown by the client if so desired.
 * <p>
 * Of the Workbook implementation classes, both XSSFWorkbook (the default .xlsx Workbook) and SXSSFWorkbook (the streaming variant) are supported.
 * Both of the experimental sub-classes of SXSSFWorkbook will probably work fine, but no guarantees can be given.
 * The HSSFWorkbook implementation class is NOT supported, since that class is used for writing .xls files.
 */
public class SkinnyFileWriter implements XlsxFileWriter {
	private static final Logger LOG = LoggerFactory.getLogger(SkinnyFileWriter.class);
	private static final String PATTERN = "yyyy-MM-dd-HH-mm-ss-SSS";
	private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern(PATTERN);

	private boolean lastWriteSuccessful;
	private Exception lastWriteException;

	@Override
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

	@Override
	public boolean isLastWriteSuccessful() {
		return lastWriteSuccessful;
	}

	@Override
	public Optional<Exception> getLastWriteException() {
		return Optional.ofNullable(lastWriteException);
	}

	private void writeContentToFileSystem(Workbook content, File outputFile) throws IOException {
		outputFile.createNewFile();

		try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
			content.write(outputStream);
		}

	}

	private static File sanitizeOutputFile(File outputFile) throws InterruptedException {
		if (!outputFile.getParentFile().exists()) {
			LOG.info("Directory {} does not exist and will be created before writing to file {}",
					outputFile.getParentFile(), outputFile);
			outputFile.getParentFile().mkdirs();
		}
		if (outputFile.exists()) {
			File actualOutputFile = getActualOutputFile(outputFile);
			if (actualOutputFile.exists()) {
				Thread.sleep(1);
				actualOutputFile = getActualOutputFile(outputFile);
			}

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
		return LocalDateTime.now().format(FORMATTER);
	}

}
