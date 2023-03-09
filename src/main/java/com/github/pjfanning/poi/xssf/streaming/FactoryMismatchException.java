package com.github.pjfanning.poi.xssf.streaming;

/**
 * An exception that is thrown when a {@link CommentsTableFactory} or {@link SharedStringsTableFactory}
 * return an instance that is not a valid {@link org.apache.poi.ooxml.POIXMLDocument}.
 */
public class FactoryMismatchException extends RuntimeException {
    public FactoryMismatchException(String message) {
        super(message);
    }
}
