package org.rdlinux.xlsx.impl;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.rdlinux.xlsx.exceptions.CloseException;
import org.rdlinux.xlsx.exceptions.ParseException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.*;
import java.util.*;

public class StreamingSheetReader implements Iterable<Row> {
    private static final Logger log = LoggerFactory.getLogger(StreamingSheetReader.class);

    private final SharedStringsTable sst;
    private final StylesTable stylesTable;
    private final XMLEventReader parser;
    private final DataFormatter dataFormatter = new DataFormatter();
    private final Set<Integer> hiddenColumns = new HashSet<>();

    private int lastRowNum;
    private int currentRowNum;
    private int firstColNum = 0;
    private int currentColNum;
    private int rowCacheSize;
    private List<Row> rowCache = new ArrayList<>();
    private Iterator<Row> rowCacheIterator;

    private String lastContents;
    private Sheet sheet;
    private StreamingRow currentRow;
    private StreamingCell currentCell;
    private boolean use1904Dates;

    public StreamingSheetReader(SharedStringsTable sst, StylesTable stylesTable, XMLEventReader parser,
                                final boolean use1904Dates, int rowCacheSize) {
        this.sst = sst;
        this.stylesTable = stylesTable;
        this.parser = parser;
        this.use1904Dates = use1904Dates;
        this.rowCacheSize = rowCacheSize;
    }

    void setSheet(StreamingSheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Read through a number of rows equal to the rowCacheSize field or until there is no more data to read
     *
     * @return true if data was read
     */
    private boolean getRow() {
        try {
            this.rowCache.clear();
            while (this.rowCache.size() < this.rowCacheSize && this.parser.hasNext()) {
                this.handleEvent(this.parser.nextEvent());
            }
            this.rowCacheIterator = this.rowCache.iterator();
            return this.rowCacheIterator.hasNext();
        } catch (XMLStreamException e) {
            throw new ParseException("Error reading XML stream", e);
        }
    }

    private String[] splitCellRef(String ref) {
        int splitPos = -1;

        // start at pos 1, since the first char is expected to always be a letter
        for (int i = 1; i < ref.length(); i++) {
            char c = ref.charAt(i);

            if (c >= '0' && c <= '9') {
                splitPos = i;
                break;
            }
        }

        return new String[]{
                ref.substring(0, splitPos),
                ref.substring(splitPos)
        };
    }

    /**
     * Handles a SAX event.
     *
     * @param event
     */
    private void handleEvent(XMLEvent event) {
        if (event.getEventType() == XMLStreamConstants.CHARACTERS) {
            Characters c = event.asCharacters();
            this.lastContents += c.getData();
        } else if (event.getEventType() == XMLStreamConstants.START_ELEMENT
                && this.isSpreadsheetTag(event.asStartElement().getName())) {
            StartElement startElement = event.asStartElement();
            String tagLocalName = startElement.getName().getLocalPart();

            if ("row".equals(tagLocalName)) {
                Attribute rowNumAttr = startElement.getAttributeByName(new QName("r"));
                int rowIndex = this.currentRowNum;
                if (rowNumAttr != null) {
                    rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
                    this.currentRowNum = rowIndex;
                }
                Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
                boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
                this.currentRow = new StreamingRow(this.sheet, rowIndex, isHidden);
                this.currentColNum = this.firstColNum;
            } else if ("col".equals(tagLocalName)) {
                Attribute isHiddenAttr = startElement.getAttributeByName(new QName("hidden"));
                boolean isHidden = isHiddenAttr != null && ("1".equals(isHiddenAttr.getValue()) || "true".equals(isHiddenAttr.getValue()));
                if (isHidden) {
                    Attribute minAttr = startElement.getAttributeByName(new QName("min"));
                    Attribute maxAttr = startElement.getAttributeByName(new QName("max"));
                    int min = Integer.parseInt(minAttr.getValue()) - 1;
                    int max = Integer.parseInt(maxAttr.getValue()) - 1;
                    for (int columnIndex = min; columnIndex <= max; columnIndex++) {
                        this.hiddenColumns.add(columnIndex);
                    }
                }
            } else if ("c".equals(tagLocalName)) {
                Attribute ref = startElement.getAttributeByName(new QName("r"));

                if (ref != null) {
                    String[] coord = this.splitCellRef(ref.getValue());
                    this.currentColNum = CellReference.convertColStringToIndex(coord[0]);
                    this.currentCell = new StreamingCell(this.sheet, this.currentColNum, Integer.parseInt(coord[1]) - 1, this.use1904Dates);
                } else {
                    this.currentCell = new StreamingCell(this.sheet, this.currentColNum, this.currentRowNum, this.use1904Dates);
                }
                this.setFormatString(startElement, this.currentCell);

                Attribute type = startElement.getAttributeByName(new QName("t"));
                if (type != null) {
                    this.currentCell.setType(type.getValue());
                } else {
                    this.currentCell.setType("n");
                }

                Attribute style = startElement.getAttributeByName(new QName("s"));
                if (style != null) {
                    String indexStr = style.getValue();
                    try {
                        int index = Integer.parseInt(indexStr);
                        this.currentCell.setCellStyle(this.stylesTable.getStyleAt(index));
                    } catch (NumberFormatException nfe) {
                        log.warn("Ignoring invalid style index {}", indexStr);
                    }
                } else {
                    this.currentCell.setCellStyle(this.stylesTable.getStyleAt(0));
                }
            } else if ("dimension".equals(tagLocalName)) {
                Attribute refAttr = startElement.getAttributeByName(new QName("ref"));
                String ref = refAttr != null ? refAttr.getValue() : null;
                if (ref != null) {
                    // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
                    for (int i = ref.length() - 1; i >= 0; i--) {
                        if (!Character.isDigit(ref.charAt(i))) {
                            try {
                                this.lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
                            } catch (NumberFormatException ignore) {
                            }
                            break;
                        }
                    }
                    for (int i = 0; i < ref.length(); i++) {
                        if (!Character.isAlphabetic(ref.charAt(i))) {
                            this.firstColNum = CellReference.convertColStringToIndex(ref.substring(0, i));
                            break;
                        }
                    }
                }
            } else if ("f".equals(tagLocalName)) {
                if (this.currentCell != null) {
                    this.currentCell.setFormulaType(true);
                }
            }

            // Clear contents cache
            this.lastContents = "";
        } else if (event.getEventType() == XMLStreamConstants.END_ELEMENT
                && this.isSpreadsheetTag(event.asEndElement().getName())) {
            EndElement endElement = event.asEndElement();
            String tagLocalName = endElement.getName().getLocalPart();

            if ("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
                this.currentCell.setRawContents(this.unformattedContents());
                this.currentCell.setContentSupplier(this.formattedContents());
            } else if ("row".equals(tagLocalName) && this.currentRow != null) {
                this.rowCache.add(this.currentRow);
                this.currentRowNum++;
            } else if ("c".equals(tagLocalName)) {
                this.currentRow.getCellMap().put(this.currentCell.getColumnIndex(), this.currentCell);
                this.currentCell = null;
                this.currentColNum++;
            } else if ("f".equals(tagLocalName)) {
                if (this.currentCell != null) {
                    this.currentCell.setFormula(this.lastContents);
                }
            }

        }
    }

    /**
     * Returns true if a tag is part of the main namespace for SpreadsheetML:
     * <ul>
     * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
     * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
     * </ul>
     * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
     *
     * @param name
     * @return
     */
    private boolean isSpreadsheetTag(QName name) {
        return (name.getNamespaceURI() != null
                && name.getNamespaceURI().endsWith("/main"));
    }

    /**
     * Get the hidden state for a given column
     *
     * @param columnIndex - the column to set (0-based)
     * @return hidden - <code>false</code> if the column is visible
     */
    boolean isColumnHidden(int columnIndex) {
        if (this.rowCacheIterator == null) {
            this.getRow();
        }
        return this.hiddenColumns.contains(columnIndex);
    }

    /**
     * Gets the last row on the sheet
     *
     * @return
     */
    int getLastRowNum() {
        if (this.rowCacheIterator == null) {
            this.getRow();
        }
        return this.lastRowNum;
    }

    /**
     * Read the numeric format string out of the styles table for this cell. Stores
     * the result in the Cell.
     *
     * @param startElement
     * @param cell
     */
    void setFormatString(StartElement startElement, StreamingCell cell) {
        Attribute cellStyle = startElement.getAttributeByName(new QName("s"));
        String cellStyleString = (cellStyle != null) ? cellStyle.getValue() : null;
        XSSFCellStyle style = null;

        if (cellStyleString != null) {
            style = this.stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
        } else if (this.stylesTable.getNumCellStyles() > 0) {
            style = this.stylesTable.getStyleAt(0);
        }

        if (style != null) {
            cell.setNumericFormatIndex(style.getDataFormat());
            String formatString = style.getDataFormatString();

            if (formatString != null) {
                cell.setNumericFormat(formatString);
            } else {
                cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
            }
        } else {
            cell.setNumericFormatIndex(null);
            cell.setNumericFormat(null);
        }
    }

    /**
     * Tries to format the contents of the last contents appropriately based on
     * the type of cell and the discovered numeric format.
     *
     * @return
     */
    Supplier formattedContents() {
        return this.getFormatterForType(this.currentCell.getType());
    }

    /**
     * Tries to format the contents of the last contents appropriately based on
     * the provided type and the discovered numeric format.
     *
     * @return
     */
    private Supplier getFormatterForType(String type) {
        switch (type) {
            case "s":           //string stored in shared table
                if (!this.lastContents.isEmpty()) {
                    int idx = Integer.parseInt(this.lastContents);
                    return new StringSupplier(this.sst.getItemAt(idx).toString());
                }
                return new StringSupplier(this.lastContents);
            case "inlineStr":   //inline string (not in sst)
            case "str":
                return new StringSupplier(new XSSFRichTextString(this.lastContents).toString());
            case "e":           //error type
                return new StringSupplier("ERROR:  " + this.lastContents);
            case "n":           //numeric type
                if (this.currentCell.getNumericFormat() != null && this.lastContents.length() > 0) {
                    // the formatRawCellContents operation incurs a significant overhead on large sheets,
                    // and we want to defer the execution of this method until the value is actually needed.
                    // it is not needed in all cases..
                    final String currentLastContents = this.lastContents;
                    final int currentNumericFormatIndex = this.currentCell.getNumericFormatIndex();
                    final String currentNumericFormat = this.currentCell.getNumericFormat();

                    return new Supplier() {
                        String cachedContent;

                        @Override
                        public Object getContent() {
                            if (this.cachedContent == null) {
                                this.cachedContent = StreamingSheetReader.this.dataFormatter.formatRawCellContents(
                                        Double.parseDouble(currentLastContents),
                                        currentNumericFormatIndex,
                                        currentNumericFormat);
                            }

                            return this.cachedContent;
                        }
                    };
                } else {
                    return new StringSupplier(this.lastContents);
                }
            default:
                return new StringSupplier(this.lastContents);
        }
    }

    /**
     * Returns the contents of the cell, with no formatting applied
     *
     * @return
     */
    String unformattedContents() {
        switch (this.currentCell.getType()) {
            case "s":           //string stored in shared table
                if (!this.lastContents.isEmpty()) {
                    int idx = Integer.parseInt(this.lastContents);
                    return this.sst.getItemAt(idx).toString();
                }
                return this.lastContents;
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(this.lastContents).toString();
            default:
                return this.lastContents;
        }
    }

    /**
     * Returns a new streaming iterator to loop through rows. This iterator is not
     * guaranteed to have all rows in memory, and any particular iteration may
     * trigger a load from disk to read in new data.
     *
     * @return the streaming iterator
     */
    @Override
    public Iterator<Row> iterator() {
        return new StreamingRowIterator();
    }

    public void close() {
        try {
            this.parser.close();
        } catch (XMLStreamException e) {
            throw new CloseException(e);
        }
    }

    class StreamingRowIterator implements Iterator<Row> {
        public StreamingRowIterator() {
            if (StreamingSheetReader.this.rowCacheIterator == null) {
                this.hasNext();
            }
        }

        @Override
        public boolean hasNext() {
            return (StreamingSheetReader.this.rowCacheIterator != null && StreamingSheetReader.this.rowCacheIterator.hasNext()) || StreamingSheetReader.this.getRow();
        }

        @Override
        public Row next() {
            return StreamingSheetReader.this.rowCacheIterator.next();
        }

        @Override
        public void remove() {
            throw new RuntimeException("NotSupported");
        }
    }
}
