package com.github.neutius.skinny.xlsx.writer.interfaces;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.Optional;

/**
 * Any implementation of this interface writes the content of an in memory Apache POI Workbook to a file on the file system.
 * <p>
 * After any write action, feedback on the success of that write action is available in the form of a boolean.
 * Any Exception that was caught during that write action is also available, and can be rethrown by the client if so desired.
 */
public interface XlsxFileWriter {

	/**
	 * Write the content of a Workbook to a File on the file system.
	 * Any exception that occurs wil be caught, and will remain available for retrieval until this method is called again.
	 *
	 * @param content    The content to be written to the file system.
	 * @param outputFile The location of the file to be written.
	 *                   If this file already exists, a new file within the same directory will be created.
	 *                   If this file is in a directory that does not exist, that directory and any non-existent parent directories will be created.
	 */
	void write(Workbook content, File outputFile);

	/**
	 * Provides feedback on the success of the previous write attempt.
	 * <p>
	 * Returns true if no Exception occurred during the previous write action.
	 * Returns false if an Exception did occur, or if no write action has been performed.
	 *
	 * @return A best-effort indication of whether your content has been written to the file system.
	 */
	boolean isLastWriteSuccessful();

	/**
	 * Any Exception that occurs during a write action is stored and available for later retrieval, until another write action is performed.
	 *
	 * @return the Exception that may or may not have occurred during the previous write action.
	 */
	Optional<Exception> getLastWriteException();
}
