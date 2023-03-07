package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Can be used with {@link org.apache.poi.xssf.usermodel.XSSFWorkbook} or {@link org.apache.poi.xssf.streaming.SXSSFWorkbook}
 * constructors to override the implementation of {@link org.apache.poi.xssf.model.Comments} and/or
 * {@link org.apache.poi.xssf.model.SharedStrings}. Use {@link #builder()} to build your factory instance.
 * You should create a new instance of the factory for each <code>Workbook</code> that you work with. This is because
 * the {@link org.apache.poi.xssf.model.CommentsTable} and {@link org.apache.poi.xssf.model.SharedStringsTable}
 * are not reusable.
 *
 * @see SXSSFFactory an alternative factory class that is easier to use but limited in the type of table implememtations it supports
 * @see TempFileSharedStringsTable
 * @see TempFileCommentsTable
 * @see MapBackedSharedStringsTable
 * @see MapBackedCommentsTable
 * @see SharedStringsTable
 * @see CommentsTable
 */
public class CustomXSSFFactory extends XSSFFactory {

    public static class Builder {
        private Comments commentsTable;
        private SharedStrings sharedStrings;

        /**
         * @param commentsTable a comment table instance
         * @return this Builder instance
         * @throws ConfigException if the <code>Comments</code> instance is not a subclass of {@link POIXMLDocumentPart}
         */
        public Builder commentsTable(Comments commentsTable) throws ConfigException {
            if (commentsTable instanceof POIXMLDocumentPart) {
                this.commentsTable = commentsTable;
            } else {
                throw new ConfigException("Comments instance must be a subclass of POIXMLDocumentPart");
            }
            return this;
        }

        /**
         * @param sharedStrings a sharedStrings table instance
         * @return this Builder instance
         * @throws ConfigException if the <code>SharedStrings</code> instance is not a subclass of {@link POIXMLDocumentPart}
         */
        public Builder sharedStrings(SharedStrings sharedStrings) throws ConfigException {
            if (sharedStrings instanceof POIXMLDocumentPart) {
                this.sharedStrings = sharedStrings;
            } else {
                throw new ConfigException("SharedStrings instance must be a subclass of POIXMLDocumentPart");
            }
            this.sharedStrings = sharedStrings;
            return this;
        }

        public CustomXSSFFactory build() {
            return new CustomXSSFFactory(commentsTable, sharedStrings);
        }
    }

    public static Builder builder() {
        return new Builder();
    }

    private final Comments commentsTable;
    private final SharedStrings sharedStrings;


    private CustomXSSFFactory(Comments commentsTable, SharedStrings sharedStrings) {
        this.commentsTable = commentsTable;
        this.sharedStrings = sharedStrings;
    }

    @Override
    public POIXMLDocumentPart newDocumentPart(POIXMLRelation descriptor) {
        if (XSSFRelation.SHARED_STRINGS.getRelation().equals(descriptor.getRelation()) &&
                sharedStrings instanceof POIXMLDocumentPart) {
            return (POIXMLDocumentPart) sharedStrings;
        }
        if (XSSFRelation.SHEET_COMMENTS.getRelation().equals(descriptor.getRelation()) &&
                commentsTable instanceof POIXMLDocumentPart) {
            return (POIXMLDocumentPart) commentsTable;
        }
        return super.newDocumentPart(descriptor);
    }
}
