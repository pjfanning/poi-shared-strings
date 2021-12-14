package com.github.pjfanning.poi.xssf.streaming.cache.lru;

import java.util.Iterator;
import java.util.LinkedHashMap;

/**
 * <p>
 *     This very simple LRUCache implementation doesn't support concurrent operations. It has been designed for
 *     processing a file in a single thread.
 * </p>
 * @param <K>
 * @param <V>
 */
public class LRUCache<K, V> {

    private final long capacity;
    private final LinkedHashMap<K, V> map = new LinkedHashMap<>();

    LRUCache(long capacity) {
        this.capacity = capacity;
    }

    V get(K key) {
        V value = map.get(key);
        if (value != null) {
            map.remove(key);
            map.put(key, value);
        }
        return value;
    }

    void put(K key, V val) {
        Iterator<V> it = map.values().iterator();
        if (map.size() >= capacity) {
            it.next();
            it.remove();
        }
        map.put(key, val);
    }
}
