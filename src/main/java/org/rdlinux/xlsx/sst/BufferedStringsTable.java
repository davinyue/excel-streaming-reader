package org.rdlinux.xlsx.sst;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class BufferedStringsTable extends SharedStringsTable implements AutoCloseable {
    private final FileBackedList list;

    public static BufferedStringsTable getSharedStringsTable(File tmp, int cacheSizeBytes, OPCPackage pkg)
            throws IOException {
        List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
        return parts.size() == 0 ? null : new BufferedStringsTable(parts.get(0), tmp, cacheSizeBytes);
    }

    private BufferedStringsTable(PackagePart part, File file, int cacheSizeBytes) throws IOException {
        this.list = new FileBackedList(file, cacheSizeBytes);
        this.readFrom(part.getInputStream());
    }

    @Override
    public void readFrom(InputStream is) throws IOException {
        try {
            XMLEventReader xmlEventReader = StaxHelper.newXMLInputFactory().createXMLEventReader(is);

            while (xmlEventReader.hasNext()) {
                XMLEvent xmlEvent = xmlEventReader.nextEvent();

                if (xmlEvent.isStartElement() && xmlEvent.asStartElement().getName().getLocalPart().equals("si")) {
                    this.list.add(this.parseCT_Rst(xmlEventReader));
                }
            }
        } catch (XMLStreamException e) {
            throw new IOException(e);
        }
    }

    /**
     * Parses a {@code <si>} String Item. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sharedstringitem.aspx">xmlschema
     * type {@code CT_Rst}</a>.
     */
    private String parseCT_Rst(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <si>;  Post condition: pointing to </si>
        StringBuilder buf = new StringBuilder();
        XMLEvent xmlEvent;
        while ((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch (xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "r": // Rich Text Run
                    this.parseCT_RElt(xmlEventReader, buf);
                    break;
                case "rPh": // Phonetic Run
                case "phoneticPr": // Phonetic Properties
                    this.skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return buf.toString();
    }

    /**
     * Parses a {@code <r>} Rich Text Run. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.run.aspx">xmlschema
     * type {@code CT_RElt}</a>.
     */
    private void parseCT_RElt(XMLEventReader xmlEventReader, StringBuilder buf) throws XMLStreamException {
        // Precondition: pointing to <r>;  Post condition: pointing to </r>
        XMLEvent xmlEvent;
        while ((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch (xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "rPr": // Run Properties
                    this.skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
    }

    private void skipElement(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to start element;  Post condition: pointing to end element
        while (xmlEventReader.nextTag().isStartElement()) {
            this.skipElement(xmlEventReader); // recursively skip over child
        }
    }

    @Override
    public RichTextString getItemAt(int idx) {
        return new XSSFRichTextString(this.list.getAt(idx));
    }

    @Override
    public void close() throws IOException {
        super.close();
        this.list.close();
    }
}
