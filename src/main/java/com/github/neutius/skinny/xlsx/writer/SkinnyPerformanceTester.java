package com.github.neutius.skinny.xlsx.writer;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.UUID;

/**
 * This class has been used to test performance between different implementations and options.
 *
 * Results with 10 sheets, 10 columns, 100000 rows (for a total of 10 million fields):
 * - SkinnyStreamer, no auto adjusting of column size: between 37 and 39 seconds.
 * - SkinnyStreamer, auto adjusting of column size based on column headers only: between 37 and 39 seconds.
 * - SkinnyStreamer, auto adjusting of column size based on last 100 rows only: between 37 and 39 seconds.
 * - SkinnyStreamer, auto adjusting of column size based on all rows: 13 minutes and between 11 and 15 seconds.
 * - SkinnyWriter: Exception in thread "main" java.lang.OutOfMemoryError: Java heap space
 */

class SkinnyPerformanceTester {
    private static final List<SkinnySheetContent> sheetContentList = createLargeSheets();

    private static final int sheetAmount = 10;
    private static final int columnAmount = 10;
    private static final int rowAmount = 100000;

    private static int counter = 1;

    public static void main(String... args) throws IOException {
        File targetFolder = Files.createTempDirectory("skinny_").toFile();

        SkinnyPerformanceTester tester = new SkinnyPerformanceTester();

        tester.skinnyStreamer_currentVersion_noAutoAdjust(targetFolder);
        tester.skinnyStreamer_currentVersion_noAutoAdjust(targetFolder);
        tester.skinnyStreamer_currentVersion_noAutoAdjust(targetFolder);

        tester.skinnyStreamer_newVersion_adjustForHeadersOnly(targetFolder);
        tester.skinnyStreamer_newVersion_adjustForHeadersOnly(targetFolder);
        tester.skinnyStreamer_newVersion_adjustForHeadersOnly(targetFolder);

        tester.skinnyStreamer_newVersion_adjustForLast100RowsOnly(targetFolder);
        tester.skinnyStreamer_newVersion_adjustForLast100RowsOnly(targetFolder);
        tester.skinnyStreamer_newVersion_adjustForLast100RowsOnly(targetFolder);

        tester.skinnyStreamer_currentVersion_withAutoAdjust(targetFolder);
        tester.skinnyStreamer_currentVersion_withAutoAdjust(targetFolder);
        tester.skinnyStreamer_currentVersion_withAutoAdjust(targetFolder);

        tester.skinnyWriter_currentVersion_hasAutoAdjust(targetFolder);
        tester.skinnyWriter_currentVersion_hasAutoAdjust(targetFolder);
        tester.skinnyWriter_currentVersion_hasAutoAdjust(targetFolder);

        deleteTempFilesRecursively(targetFolder);

    }

    private static void deleteTempFilesRecursively(File file) {
        if (file.isFile()) {
            deleteFile(file);
        }
        if (file.isDirectory()) {
            Arrays.asList(file.listFiles()).forEach(SkinnyPerformanceTester::deleteTempFilesRecursively);
            deleteFile(file);
        }
    }

    private static void deleteFile(File file) {
        boolean deleted = file.delete();
        System.out.println("Temporary file " + file.toString() + " is deleted? " + deleted);
    }

    void skinnyWriter_currentVersion_hasAutoAdjust(File targetFolder) throws IOException {
        String methodName = "skinnyWriter_currentVersion_hasAutoAdjust---------";

        Instant start = Instant.now();
        SkinnyWriter.writeContentToFileSystem(targetFolder, methodName + counter++, sheetContentList);
        Instant end = Instant.now();

        Duration timeElapsed = Duration.between(start, end);
        System.out.println(methodName + " - Time Elapsed: " + timeElapsed);
    }

    void skinnyStreamer_currentVersion_noAutoAdjust(File targetFolder) throws IOException {
        String methodName = "skinnyStreamer_currentVersion_noAutoAdjust--------";

        Instant start = Instant.now();
//        SkinnyStreamer.writeContentToFileSystem(targetFolder, methodName + counter++, sheetContentList, false);
        Instant end = Instant.now();

        Duration timeElapsed = Duration.between(start, end);
        System.out.println(methodName + " - Time Elapsed: " + timeElapsed);
    }

    void skinnyStreamer_currentVersion_withAutoAdjust(File targetFolder) throws IOException {
        String methodName = "skinnyStreamer_currentVersion_withAutoAdjust------";

        Instant start = Instant.now();
//        SkinnyStreamer.writeContentToFileSystem(targetFolder, methodName + counter++, sheetContentList, true);
        Instant end = Instant.now();

        Duration timeElapsed = Duration.between(start, end);
        System.out.println(methodName + " - Time Elapsed: " + timeElapsed);
    }

    void skinnyStreamer_newVersion_adjustForHeadersOnly(File targetFolder) throws IOException {
        String methodName = "skinnyStreamer_newVersion_adjustForHeadersOnly----";

        Instant start = Instant.now();
//        SkinnyStreamer.writeContentToFileSystem_adjustForHeadersOnly(targetFolder, methodName + counter++, sheetContentList, true);
        Instant end = Instant.now();

        Duration timeElapsed = Duration.between(start, end);
        System.out.println(methodName + " - Time Elapsed: " + timeElapsed);
    }

    private void skinnyStreamer_newVersion_adjustForLast100RowsOnly(File targetFolder) throws IOException {
        String methodName = "skinnyStreamer_newVersion_adjustForLast100RowsOnly";

        Instant start = Instant.now();
//        SkinnyStreamer.writeContentToFileSystem_adjustForLast100RowsOnly(targetFolder, methodName + counter++, sheetContentList, true);
        Instant end = Instant.now();

        Duration timeElapsed = Duration.between(start, end);
        System.out.println(methodName + " - Time Elapsed: " + timeElapsed);

    }

    private static List<SkinnySheetContent> createLargeSheets() {
        Instant start = Instant.now();

        List<List<String>> contentRows = getSheetContentWithRandomStrings();
        List<String> columnHeaders = getListOfRandomStrings();

        List<SkinnySheetContent> result = new ArrayList<>();

        for (int i = 0; i < sheetAmount; i++) {
            result.add(DefaultSheetContent.withHeaders("sheet" + i, columnHeaders, contentRows));
        }

        Instant end = Instant.now();
        Duration timeElapsed = Duration.between(start, end);
        System.out.println("Created sheet contents - Time Elapsed: " + timeElapsed);

        return result;
    }

    private static List<List<String>> getSheetContentWithRandomStrings() {
        List<List<String>> result = new ArrayList<>();

        for (int i = 0; i < rowAmount; i++) {
            result.add(getListOfRandomStrings());
        }

        return result;
    }

    private static List<String> getListOfRandomStrings() {
        List<String> result = new ArrayList<>();

        for (int i = 0; i < columnAmount; i++) {
            result.add(UUID.randomUUID().toString());
        }

        return result;
    }


}
