package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;

public class SXSSFFactory extends XSSFFactory {

    private boolean encryptTempFiles = false;
    private boolean enableTempFileComments = false;

    public SXSSFFactory() {}

    @Deprecated
    public SXSSFFactory(boolean encryptTempFiles) {
        super();
        this.encryptTempFiles = encryptTempFiles;
    }

    public SXSSFFactory encryptTempFiles(boolean encryptTempFiles) {
        this.encryptTempFiles = encryptTempFiles;
        return this;

    }
    public SXSSFFactory enableTempFileComments(boolean enableTempFileComments) {
        this.enableTempFileComments = enableTempFileComments;
        return this;
    }

    @Override
    public POIXMLDocumentPart newDocumentPart(POIXMLRelation descriptor) {
        if (XSSFRelation.SHARED_STRINGS.getRelation().equals(descriptor.getRelation())) {
            try {
                return new TempFileSharedStringsTable(encryptTempFiles);
            } catch (Error|RuntimeException e) {
                throw new RuntimeException("Exception creating TempFileSharedStringsTable; com.h2database h2 jar is " +
                        "required for this feature and is not included as a core dependency of poi-ooxml");
            }
        }
        if (XSSFRelation.SHEET_COMMENTS.getRelation().equals(descriptor.getRelation()) && enableTempFileComments) {
            try {
                return new TempFileCommentsTable(encryptTempFiles);
            } catch (Error|RuntimeException e) {
                throw new RuntimeException("Exception creating TempFileCommentsTable; com.h2database h2 jar is " +
                        "required for this feature and is not included as a core dependency of poi-ooxml");
            }
        }
        return super.newDocumentPart(descriptor);
    }
}
