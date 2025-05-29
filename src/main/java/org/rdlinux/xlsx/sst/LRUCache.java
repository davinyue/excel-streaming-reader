package org.rdlinux.xlsx.sst;

import java.util.Iterator;
import java.util.LinkedHashMap;

class LRUCache {

    private long sizeBytes;
    private final long capacityBytes;
    private final LinkedHashMap<Integer, String> map = new LinkedHashMap<>();

    LRUCache(long capacityBytes) {
        this.capacityBytes = capacityBytes;
    }

    String getIfPresent(int key) {
        String s = this.map.get(key);
        if (s != null) {
            this.map.remove(key);
            this.map.put(key, s);
        }
        return s;
    }

    void store(int key, String val) {
        long valSize = strSize(val);
        if (valSize > this.capacityBytes) {
            throw new RuntimeException("Insufficient cache space.");
        }
        Iterator<String> it = this.map.values().iterator();
        while (valSize + this.sizeBytes > this.capacityBytes) {
            String s = it.next();
            this.sizeBytes -= strSize(s);
            it.remove();
        }
        this.map.put(key, val);
        this.sizeBytes += valSize;
    }

    //  just an estimation
    private static long strSize(String str) {
        long size = Integer.BYTES; // hashCode
        size += Character.BYTES * str.length(); // characters
        return size;
    }

}
