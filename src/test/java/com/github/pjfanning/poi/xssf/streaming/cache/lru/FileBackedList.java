package com.github.pjfanning.poi.xssf.streaming.cache.lru;

import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * File-backed list-like class. Allows addition of arbitrary
 * numbers of array entries (serialized to JSON) in a binary
 * packed file. Reading of entries is done with an NIO
 * channel that seeks to the entry in the file.
 * <p>
 * File entry format:
 * <ul>
 * <li>4 bytes: length of entry</li>
 * <li><i>length</i> bytes: JSON string containing the entry data</li>
 * </ul>
 * <p>
 * Pointers to the offset of each entry are kept in a {@code List<Long>}.
 * The values loaded from the the file are cached up to a maximum of
 * {@code cacheSize}. Items are evicted from the cache with an LRU algorithm.
 */
public class FileBackedList implements AutoCloseable {

    private final List<Long> pointers = new ArrayList<>();
    private final RandomAccessFile raf;
    private final FileChannel channel;
    private final LRUCache<Integer, String> cache;
    private long filesize;

    public FileBackedList(File file, final int capacity) throws IOException {
        this.raf = new RandomAccessFile(file, "rw");
        this.channel = raf.getChannel();
        this.filesize = raf.length();
        this.cache = new LRUCache<>(capacity);
    }

    public void add(String str) {
        try {
            writeToFile(str);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public String getAt(int index) {
        String s = cache.get(index);
        if (s != null) {
            return s;
        }
        try {
            String value = readFromFile(pointers.get(index));
            cache.put(index, value);
            return value;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public Integer size() {
        return pointers.size();
    }

    private void writeToFile(String str) throws IOException {
        synchronized (channel) {
            ByteBuffer bytes = ByteBuffer.wrap(str.getBytes(StandardCharsets.UTF_8));
            ByteBuffer length = ByteBuffer.allocate(4).putInt(bytes.array().length);

            channel.position(filesize);
            pointers.add(channel.position());
            length.flip();
            channel.write(length);
            channel.write(bytes);

            filesize += 4 + bytes.array().length;
        }
    }

    private String readFromFile(long pointer) throws IOException {
        synchronized (channel) {
            FileChannel fc = channel.position(pointer);

            //get length of entry
            ByteBuffer buffer = ByteBuffer.wrap(new byte[4]);
            fc.read(buffer);
            buffer.flip();
            int length = buffer.getInt();

            //read entry
            buffer = ByteBuffer.wrap(new byte[length]);
            fc.read(buffer);
            buffer.flip();

            return new String(buffer.array(), StandardCharsets.UTF_8);
        }
    }

    @Override
    public void close() {
        try {
            raf.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}