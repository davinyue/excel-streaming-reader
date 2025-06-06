package org.rdlinux.xlsx.impl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.rdlinux.xlsx.StreamingReader.Builder;
import org.rdlinux.xlsx.exceptions.OpenException;
import org.rdlinux.xlsx.exceptions.ReadException;
import org.rdlinux.xlsx.sst.BufferedStringsTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.nio.file.Files;
import java.security.GeneralSecurityException;
import java.util.*;

import static java.util.Arrays.asList;
import static org.rdlinux.xlsx.XmlUtils.document;
import static org.rdlinux.xlsx.XmlUtils.searchForNodeList;
import static org.rdlinux.xlsx.impl.TempFileUtil.writeInputStreamToFile;

public class StreamingWorkbookReader implements Iterable<Sheet>, AutoCloseable {
    private static final Logger log = LoggerFactory.getLogger(StreamingWorkbookReader.class);

    private final List<StreamingSheet> sheets;
    private final List<Map<String, String>> sheetProperties = new ArrayList<>();
    private final Builder builder;
    private File tmp;
    private File sstCache;
    private OPCPackage pkg;
    private SharedStringsTable sst;
    private boolean use1904Dates = false;

    /**
     * This constructor exists only so the StreamingReader can instantiate
     * a StreamingWorkbook using its own reader implementation. Do not use
     * going forward.
     *
     * @param sst      The SST data for this workbook
     * @param sstCache The backing cache file for the SST data
     * @param pkg      The POI package that should be closed when this workbook is closed
     * @param reader   A single streaming reader instance
     * @param builder  The builder containing all options
     */
    @Deprecated
    public StreamingWorkbookReader(SharedStringsTable sst, File sstCache, OPCPackage pkg, StreamingSheetReader reader, Builder builder) {
        this.sst = sst;
        this.sstCache = sstCache;
        this.pkg = pkg;
        this.sheets = asList(new StreamingSheet(null, reader));
        this.builder = builder;
    }

    public StreamingWorkbookReader(Builder builder) {
        this.sheets = new ArrayList<>();
        this.builder = builder;
    }

    public StreamingSheetReader first() {
        return this.sheets.get(0).getReader();
    }

    public void init(InputStream is) {
        File f = null;
        try {
            f = writeInputStreamToFile(is, this.builder.getBufferSize());
            log.debug("Created temp file [" + f.getAbsolutePath() + "]");

            this.init(f);
            this.tmp = f;
        } catch (IOException e) {
            throw new ReadException("Unable to read input stream", e);
        } catch (RuntimeException e) {
            if (f != null) {
                f.delete();
            }
            throw e;
        }
    }

    public void init(File f) {
        try {
            if (this.builder.getPassword() != null) {
                // Based on: https://poi.apache.org/encryption.html
                POIFSFileSystem poifs = new POIFSFileSystem(f);
                EncryptionInfo info = new EncryptionInfo(poifs);
                Decryptor d = Decryptor.getInstance(info);
                d.verifyPassword(this.builder.getPassword());
                this.pkg = OPCPackage.open(d.getDataStream(poifs));
            } else {
                this.pkg = OPCPackage.open(f);
            }

            XSSFReader reader = new XSSFReader(this.pkg);
            if (this.builder.getSstCacheSizeBytes() > 0) {
                this.sstCache = Files.createTempFile("", "").toFile();
                log.debug("Created sst cache file [" + this.sstCache.getAbsolutePath() + "]");
                this.sst = BufferedStringsTable.getSharedStringsTable(this.sstCache, this.builder.getSstCacheSizeBytes(), this.pkg);
            } else {
                this.sst = (SharedStringsTable) reader.getSharedStringsTable();
            }

            StylesTable styles = reader.getStylesTable();
            NodeList workbookPr = searchForNodeList(document(reader.getWorkbookData()), "/ss:workbook/ss:workbookPr");
            if (workbookPr.getLength() == 1) {
                final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
                if (date1904 != null) {
                    this.use1904Dates = ("1".equals(date1904.getTextContent()));
                }
            }

            this.loadSheets(reader, this.sst, styles, this.builder.getRowCacheSize());
        } catch (IOException e) {
            throw new OpenException("Failed to open file", e);
        } catch (OpenXML4JException | XMLStreamException e) {
            throw new ReadException("Unable to read workbook", e);
        } catch (GeneralSecurityException e) {
            throw new ReadException("Unable to read workbook - Decryption failed", e);
        }
    }

    void loadSheets(XSSFReader reader, SharedStringsTable sst, StylesTable stylesTable, int rowCacheSize)
            throws IOException, InvalidFormatException, XMLStreamException {
        this.lookupSheetNames(reader);

        //Some workbooks have multiple references to the same sheet. Need to filter
        //them out before creating the XMLEventReader by keeping track of their URIs.
        //The sheets are listed in order, so we must keep track of insertion order.
        SheetIterator iter = (SheetIterator) reader.getSheetsData();
        Map<URI, InputStream> sheetStreams = new LinkedHashMap<>();
        while (iter.hasNext()) {
            InputStream is = iter.next();
            sheetStreams.put(iter.getSheetPart().getPartName().getURI(), is);
        }

        //Iterate over the loaded streams
        int i = 0;
        for (URI uri : sheetStreams.keySet()) {
            XMLEventReader parser = StaxHelper.newXMLInputFactory().createXMLEventReader(sheetStreams.get(uri));
            this.sheets.add(new StreamingSheet(this.sheetProperties.get(i++).get("name"), new StreamingSheetReader(sst, stylesTable, parser, this.use1904Dates, rowCacheSize)));
        }
    }

    void lookupSheetNames(XSSFReader reader) throws IOException, InvalidFormatException {
        this.sheetProperties.clear();
        NodeList nl = searchForNodeList(document(reader.getWorkbookData()), "/ss:workbook/ss:sheets/ss:sheet");
        for (int i = 0; i < nl.getLength(); i++) {
            Map<String, String> props = new HashMap<>();
            props.put("name", nl.item(i).getAttributes().getNamedItem("name").getTextContent());

            Node state = nl.item(i).getAttributes().getNamedItem("state");
            props.put("state", state == null ? "visible" : state.getTextContent());
            this.sheetProperties.add(props);
        }
    }

    List<? extends Sheet> getSheets() {
        return this.sheets;
    }

    public List<Map<String, String>> getSheetProperties() {
        return this.sheetProperties;
    }

    @Override
    public Iterator<Sheet> iterator() {
        return new StreamingSheetIterator(this.sheets.iterator());
    }

    @Override
    public void close() throws IOException {
        try {
            for (StreamingSheet sheet : this.sheets) {
                sheet.getReader().close();
            }
            this.pkg.revert();
        } finally {
            if (this.tmp != null) {
                if (log.isDebugEnabled()) {
                    log.debug("Deleting tmp file [" + this.tmp.getAbsolutePath() + "]");
                }
                this.tmp.delete();
            }
            if (this.sst instanceof BufferedStringsTable) {
                if (log.isDebugEnabled()) {
                    log.debug("Deleting sst cache file [" + this.sstCache.getAbsolutePath() + "]");
                }
                ((BufferedStringsTable) this.sst).close();
                this.sstCache.delete();
            }
        }
    }

    static class StreamingSheetIterator implements Iterator<Sheet> {
        private final Iterator<StreamingSheet> iterator;

        public StreamingSheetIterator(Iterator<StreamingSheet> iterator) {
            this.iterator = iterator;
        }

        @Override
        public boolean hasNext() {
            return this.iterator.hasNext();
        }

        @Override
        public Sheet next() {
            return this.iterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }
}
