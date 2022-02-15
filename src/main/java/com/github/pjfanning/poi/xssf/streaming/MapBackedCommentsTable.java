package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Table of comments.
 * <p>
 * The comments table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public class MapBackedCommentsTable extends CommentsTableBase {
    private static Logger log = LoggerFactory.getLogger(MapBackedCommentsTable.class);

    public MapBackedCommentsTable() {
        this(false);
    }

    public MapBackedCommentsTable(boolean fullFormat) {
        super(fullFormat);
        comments = new ConcurrentHashMap<>();
        authors = new ConcurrentHashMap<>();
    }

    public MapBackedCommentsTable(OPCPackage pkg) throws IOException {
        this(pkg, false);
    }

    public MapBackedCommentsTable(OPCPackage pkg, boolean fullFormat) throws IOException {
        this(fullFormat);
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
        return authors.keySet().iterator();
    }

    @Override
    protected Iterator<String> commentsKeyIterator() {
        return comments.keySet().iterator();
    }

    @Override
    public void close() {
        comments.clear();
        authors.clear();
    }
}
