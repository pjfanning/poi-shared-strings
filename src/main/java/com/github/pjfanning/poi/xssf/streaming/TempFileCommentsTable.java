package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
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

/**
 * Table of comments.
 * <p>
 * The comments table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public class TempFileCommentsTable extends CommentsTableBase {
    private static Logger log = LoggerFactory.getLogger(TempFileCommentsTable.class);

    private File tempFile;
    private MVStore mvStore;
    private MVMap<String, SerializableComment> mvComments;
    private MVMap<Integer, String> mvAuthors;

    public TempFileCommentsTable() throws IOException {
        this(false, false);
    }

    /**
     * @param encryptTempFiles whether to encrypt the temp files
     * @throws IOException if an error occurs while working with the temp file
     */
    public TempFileCommentsTable(boolean encryptTempFiles) throws IOException {
        this(encryptTempFiles, false);
    }

    /**
     * @param encryptTempFiles whether to encrypt the temp files
     * @param fullFormat whether to store format information (which is more expensive)
     * @throws IOException if an error occurs while working with the temp file
     */
    public TempFileCommentsTable(boolean encryptTempFiles, boolean fullFormat) throws IOException {
        super(fullFormat);
        try {
            tempFile = TempFile.createTempFile("poi-comments", ".tmp");
            MVStore.Builder mvStoreBuilder = new MVStore.Builder();
            if (encryptTempFiles) {
                byte[] bytes = new byte[1024];
                Constants.RANDOM.nextBytes(bytes);
                mvStoreBuilder.encryptionKey(Base64.getEncoder().encodeToString(bytes).toCharArray());
            }
            mvStoreBuilder.fileName(tempFile.getAbsolutePath());
            mvStore = mvStoreBuilder.open();
            mvComments = mvStore.openMap("comments");
            comments = mvComments;
            mvAuthors = mvStore.openMap("authors");
            authors = mvAuthors;
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

    /**
     * @param pkg the OPCPackage to load the comments from
     * @param encryptTempFiles whether to encrypt the temp files
     * @throws IOException if an error occurs while working with the temp file
     */
    public TempFileCommentsTable(OPCPackage pkg, boolean encryptTempFiles) throws IOException {
        this(pkg, encryptTempFiles, false);
    }

    /**
     * @param pkg the OPCPackage to load the comments from
     * @param encryptTempFiles whether to encrypt the temp files
     * @param fullFormat whether to store format information (which is more expensive)
     * @throws IOException if an error occurs while working with the temp file
     */
    public TempFileCommentsTable(OPCPackage pkg, boolean encryptTempFiles,
                                 boolean fullFormat) throws IOException {
        this(encryptTempFiles, fullFormat);
        ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHEET_COMMENTS.getContentType());
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
    protected Iterator<Integer> authorsKeyIterator() {
        return mvAuthors.keyIterator(null);
    }

    @Override
    protected Iterator<String> commentsKeyIterator() {
        return mvComments.keyIterator(null);
    }

    @Override
    public void close() {
        if(mvStore != null) mvStore.closeImmediately();
        if(tempFile != null && !tempFile.delete()) {
            log.debug("failed to delete temp file - probably already deleted");
        }
    }
}
