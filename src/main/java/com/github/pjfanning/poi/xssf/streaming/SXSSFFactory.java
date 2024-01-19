package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Can be used with {@link org.apache.poi.xssf.usermodel.XSSFWorkbook} or {@link org.apache.poi.xssf.streaming.SXSSFWorkbook}
 * constructors to override the implementation of {@link org.apache.poi.xssf.model.CommentsTable} and/or
 * {@link org.apache.poi.xssf.model.SharedStringsTable}. This factory only supports {@link TempFileCommentsTable}
 * and {@link TempFileSharedStringsTable}.
 *
 * @see CustomXSSFFactory an alternative factory class that allows you more control of the injected <code>CommentsTable</code>
 * and <code>SharedStringsTable</code> implementations.
 */
public class SXSSFFactory extends XSSFFactory {

    private boolean encryptTempFiles = false;
    private boolean enableTempFileSharedStrings = true;
    private boolean enableTempFileComments = false;

    public SXSSFFactory() {}

    /**
     * @param encryptTempFiles whether to encrypt the temp files
     * @deprecated use #encryptTempFiles method instead
     */
    @Deprecated
    public SXSSFFactory(boolean encryptTempFiles) {
        super();
        this.encryptTempFiles = encryptTempFiles;
    }

    /**
     * @param encryptTempFiles whether to encrypt the temp files
     * @return this factory instance
     */
    public SXSSFFactory encryptTempFiles(boolean encryptTempFiles) {
        this.encryptTempFiles = encryptTempFiles;
        return this;
    }

    /**
     * @param enableTempFileSharedStrings whether to enable temp file shared strings table (default is true)
     * @return this factory instance
     * @since v2.2.2
     */
    public SXSSFFactory enableTempFileSharedStrings(boolean enableTempFileSharedStrings) {
        this.enableTempFileSharedStrings = enableTempFileSharedStrings;
        return this;
    }

    /**
     * @param enableTempFileComments whether to enable temp file comments table (default is false)
     * @return this factory instance
     * @since v2.3.0
     */
    public SXSSFFactory enableTempFileComments(boolean enableTempFileComments) {
        this.enableTempFileComments = enableTempFileComments;
        return this;
    }

    @Override
    public POIXMLDocumentPart newDocumentPart(POIXMLRelation descriptor) {
        if (XSSFRelation.SHARED_STRINGS.getRelation().equals(descriptor.getRelation()) && enableTempFileSharedStrings) {
            try {
                return new TempFileSharedStringsTable(encryptTempFiles);
            } catch (Exception e) {
                throw new IllegalStateException("Exception creating TempFileSharedStringsTable; com.h2database h2 jar is " +
                        "required for this feature and is not included as a core dependency of poi-shared-strings");
            }
        }
        if (XSSFRelation.SHEET_COMMENTS.getRelation().equals(descriptor.getRelation()) && enableTempFileComments) {
            try {
                return new TempFileCommentsTable(encryptTempFiles);
            } catch (Exception e) {
                throw new IllegalStateException("Exception creating TempFileCommentsTable; com.h2database h2 jar is " +
                        "required for this feature and is not included as a core dependency of poi-shared-strings");
            }
        }
        return super.newDocumentPart(descriptor);
    }
}
