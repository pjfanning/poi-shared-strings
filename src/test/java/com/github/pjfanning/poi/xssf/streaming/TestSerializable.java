package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;

import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;

import static org.junit.Assert.assertEquals;

public class TestSerializable {
    @Test
    public void testCommentFromString() throws Exception {
        SerializableComment comment = new SerializableComment();
        comment.setString("test string");
        comment.setAuthor("test author");
        comment.setAddress(new CellAddress("B20"));
        comment.setVisible(false);
        assertEquals("test string", comment.getCommentText());
        try(UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
            try(ObjectOutputStream oos = new ObjectOutputStream(bos)) {
                oos.writeObject(comment);
            }
            try(ObjectInputStream ois = new ObjectInputStream(bos.toInputStream())) {
                SerializableComment deserializedComment = (SerializableComment)ois.readObject();
                assertEquals(comment.getRow(), deserializedComment.getRow());
                assertEquals(comment.getColumn(), deserializedComment.getColumn());
                assertEquals(comment.getAddress(), deserializedComment.getAddress());
                assertEquals(comment.getString().getString(), deserializedComment.getString().getString());
                assertEquals(comment.getAuthor(), deserializedComment.getAuthor());
                assertEquals(comment.isVisible(), deserializedComment.isVisible());
                assertEquals(comment.getCommentText(), deserializedComment.getCommentText());
            }
        }
    }

    @Test
    public void testCommentFromXSSFRichTextString() throws Exception {
        SerializableComment comment = new SerializableComment();
        comment.setString(new XSSFRichTextString("test string"));
        comment.setAuthor("test author");
        comment.setAddress(new CellAddress("B20"));
        comment.setVisible(false);
        assertEquals("test string", comment.getCommentText());
        try(UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
            try(ObjectOutputStream oos = new ObjectOutputStream(bos)) {
                oos.writeObject(comment);
            }
            try(ObjectInputStream ois = new ObjectInputStream(bos.toInputStream())) {
                SerializableComment deserializedComment = (SerializableComment)ois.readObject();
                assertEquals(comment.getRow(), deserializedComment.getRow());
                assertEquals(comment.getColumn(), deserializedComment.getColumn());
                assertEquals(comment.getAddress(), deserializedComment.getAddress());
                assertEquals(comment.getString().getString(), deserializedComment.getString().getString());
                assertEquals(comment.getAuthor(), deserializedComment.getAuthor());
                assertEquals(comment.isVisible(), deserializedComment.isVisible());
                assertEquals(comment.getCommentText(), deserializedComment.getCommentText());
            }
        }
    }
}
