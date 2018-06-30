package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;

class SXSSFFactory extends XSSFFactory {

    private final boolean encryptTempFiles;

    SXSSFFactory(boolean encryptTempFiles) {
        super();
        this.encryptTempFiles = encryptTempFiles;
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
        return super.newDocumentPart(descriptor);
    }
}
