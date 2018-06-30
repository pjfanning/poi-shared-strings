package com.github.pjfanning.poi.xssf.streaming;

import java.lang.reflect.Field;
import java.security.AccessController;
import java.security.PrivilegedActionException;
import java.security.PrivilegedExceptionAction;

import static org.junit.Assert.assertTrue;

/**
 * Util class for POI JUnit TestCases, which provide additional features 
 */
public final class POITestCase {

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

}
