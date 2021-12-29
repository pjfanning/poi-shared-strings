package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.lang.reflect.Field;
import java.security.AccessController;
import java.security.PrivilegedActionException;
import java.security.PrivilegedExceptionAction;

import static org.junit.Assert.assertTrue;

/**
 * Util class for POI JUnit TestCases, which provide additional features 
 */
public final class POITestUtils {

    private POITestUtils() {}

    /**
     * Utility method to get the value of a private/protected field.
     * Only use this method in test cases!!!
     */
    public static <R,T> R getFieldValue(final Class<? super T> clazz, final T instance, final Class<R> fieldType, final String fieldName) {
        assertTrue("Reflection of private fields is only allowed for POI classes.", clazz.getName().startsWith("org.apache.poi."));
        try {
            return AccessController.doPrivileged(new PrivilegedExceptionAction<R>() {
                @Override
                public R run() throws Exception {
                    Field f = clazz.getDeclaredField(fieldName);
                    f.setAccessible(true);
                    return (R) f.get(instance);
                }
            });
        } catch (PrivilegedActionException pae) {
            throw new RuntimeException("Cannot access field '" + fieldName + "' of class " + clazz, pae.getException());
        }
    }

    public static XSSFWorkbook writeOutAndReadBack(Workbook wb) {
        // wb is usually an SXSSFWorkbook, but must also work on an XSSFWorkbook
        // since workbooks must be able to be written out and read back
        // several times in succession
        if(!(wb instanceof SXSSFWorkbook || wb instanceof XSSFWorkbook)) {
            throw new IllegalArgumentException("Expected an instance of SXSSFWorkbook");
        }

        XSSFWorkbook result;
        try {
            UnsynchronizedByteArrayOutputStream baos = new UnsynchronizedByteArrayOutputStream(8192);
            wb.write(baos);
            result = new XSSFWorkbook(baos.toInputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
