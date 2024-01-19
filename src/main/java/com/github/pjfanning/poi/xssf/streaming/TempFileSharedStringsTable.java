package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.List;

/**
 * Table of strings shared across all sheets in a workbook.
 * <p>
 * A workbook may contain thousands of cells containing string (non-numeric) data. Furthermore this data is very
 * likely to be repeated across many rows or columns. The goal of implementing a single string table that is shared
 * across the workbook is to improve performance in opening and saving the file by only reading and writing the
 * repetitive information once.
 * </p>
 * <p>
 * Consider for example a workbook summarizing information for cities within various countries. There may be a
 * column for the name of the country, a column for the name of each city in that country, and a column
 * containing the data for each city. In this case the country name is repetitive, being duplicated in many cells.
 * In many cases the repetition is extensive, and a tremendous savings is realized by making use of a shared string
 * table when saving the workbook. When displaying text in the spreadsheet, the cell table will just contain an
 * index into the string table as the value of a cell, instead of the full string.
 * </p>
 * <p>
 * The shared string table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public class TempFileSharedStringsTable extends SharedStringsTableBase {
    private static final Logger log = LoggerFactory.getLogger(TempFileSharedStringsTable.class);
    private File tempFile;
    private MVStore mvStore;
    private MVMap<Integer, String> mvStrings;

    public TempFileSharedStringsTable() throws IOException {
        this(false, false);
    }

    public TempFileSharedStringsTable(boolean encryptTempFiles) throws IOException {
        this(encryptTempFiles, false);
    }

    public TempFileSharedStringsTable(boolean encryptTempFiles, boolean fullFormat) throws IOException {
        super(fullFormat);
        try {
            tempFile = TempFile.createTempFile("poi-shared-strings", ".tmp");
            MVStore.Builder mvStoreBuilder = new MVStore.Builder();
            if (encryptTempFiles) {
                byte[] bytes = new byte[1024];
                Constants.RANDOM.nextBytes(bytes);
                mvStoreBuilder.encryptionKey(Base64.getEncoder().encodeToString(bytes).toCharArray());
            }
            mvStoreBuilder.fileName(tempFile.getAbsolutePath());
            mvStore = mvStoreBuilder.open();
            mvStrings = mvStore.openMap("strings");
            strings = mvStrings;
            stmap = mvStore.openMap("stmap");
        } catch (Error | IOException e) {
            if (mvStore != null) mvStore.closeImmediately();
            if (tempFile != null && !tempFile.delete()) {
                log.debug("failed to delete temp file - probably already deleted");
            }
            throw e;
        } catch (Exception e) {
            if (mvStore != null) mvStore.closeImmediately();
            if (tempFile != null && !tempFile.delete()) {
                log.debug("failed to delete temp file - probably already deleted");
            }
            throw new IOException(e);
        }

    }

    public TempFileSharedStringsTable(OPCPackage pkg, boolean encryptTempFiles) throws IOException {
        this(pkg, encryptTempFiles, false);
    }

    public TempFileSharedStringsTable(OPCPackage pkg, boolean encryptTempFiles,
                                      boolean fullFormat) throws IOException {
        this(encryptTempFiles, fullFormat);
        ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
        if (!parts.isEmpty()) {
            PackagePart sstPart = parts.get(0);
            this.readFrom(sstPart.getInputStream());
        }
    }

    @Override
    protected Logger getLogger() {
        return log;
    }

    @Override
    protected Iterator<Integer> keyIterator() {
        return mvStrings.keyIterator(null);
    }

    /**
     * TempFileSharedStringsTable only supports streaming access of shared strings.
     * Use <code>getItemAt</code> instead.
     *
     * @return throws UnsupportedOperationException
     * @throws UnsupportedOperationException not supported
     */
    @Override
    public List<RichTextString> getSharedStringItems() {
        throw new UnsupportedOperationException("TempFileSharedStringsTable only supports streaming access of shared strings");
    }

    @Override
    public void close() throws IOException {
        if(mvStore != null) mvStore.closeImmediately();
        if(tempFile != null && !tempFile.delete()) {
            log.debug("failed to delete temp file - probably already deleted");
        }
    }
}
