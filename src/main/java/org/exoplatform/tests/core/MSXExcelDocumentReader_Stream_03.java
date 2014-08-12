package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.security.PrivilegedExceptionAction;
import java.util.Properties;


import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.BaseDocumentReader;
import org.exoplatform.services.document.impl.POIPropertiesReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * Patched original eXo MS Excel Document Reader with the following changes .
 * <p/>
 * we only index a maximum of 5000 cells:
 * - at most 1000 cells per sheet
 * - at most the 5 first tabs of the spreadsheet
 * <p/>
 * we KEEP only the following data :
 * - cells with number (date formatted or simple number)
 * - cells with string
 * <p/>
 * we SKIP the following data :
 * - cells with blank value
 * - cells with boolean value
 * - cells with formula
 * - cells with error
 */
public class MSXExcelDocumentReader_Stream_03 extends BaseDocumentReader {

  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSXExcelDocumentReader_Stream_03");
  private static final int MAX_TABS = 5;
  private static final int MAX_CELLTAB = 1000;


  /**
   * @see org.exoplatform.services.document.DocumentReader#getMimeTypes()
   */
  public String[] getMimeTypes() {
    //Supported mimetypes:
    // "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" - "x.xlsx"
    //
    //Unsupported mimetypes:
    // "application/vnd.ms-excel.sheet.binary.macroenabled.12" - "*.xlsb"; There is exceptions at parsing
    // "application/vnd.openxmlformats-officedocument.spreadsheetml.template" - "x.xltx"; Not tested
    // "application/vnd.ms-excel.sheet.macroenabled.12" - "x.xlsm"; Not tested
    // "application/vnd.ms-excel.template.macroenabled.12" - "x.xltm"; Not tested
    // "application/vnd.ms-excel.addin.macroenabled.12" - "x.xlam"; Not tested
    return new String[]{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};
  }

  public void processSheet(
      XSSFOptimizedSheetXMLHandler_03.SheetContentsHandler sheetContentsExtractor,
      StylesTable styles,
      ReadOnlySharedStringsTable strings,
      InputStream sheetInputStream)
      throws IOException, SAXException {

    InputSource sheetSource = new InputSource(sheetInputStream);
    SAXParserFactory saxFactory = SAXParserFactory.newInstance();
    try {
      SAXParser saxParser = saxFactory.newSAXParser();
      XMLReader sheetParser = saxParser.getXMLReader();
      ContentHandler handler = new XSSFOptimizedSheetXMLHandler_03(
          styles, strings, sheetContentsExtractor, MAX_CELLTAB);
      sheetParser.setContentHandler(handler);
      sheetParser.parse(sheetSource);
    } catch (ParserConfigurationException e) {
      throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
    } catch (XSSFOptimizedSheetXMLHandler_03.StopSheetParsingException e) {
      // this exception allow us to stop the parsing of the sheet when we have reached the number of cell to parse per sheet ({@link MAX_CELLTAB }
      LOG.info(this.toString() + " - We stop parsing the sheet");
      if (LOG.isTraceEnabled()) {
        LOG.trace("We stop parsing the sheet");
      }
    }
  }

  /**
   * Returns only a text from .xlsx file content.
   *
   * @param is an input stream with .xls file content.
   * @return The string only with text from file content.
   */
  public String getContentAsText(final InputStream is) throws IOException, DocumentReadException {
    if (is == null) {
      throw new IllegalArgumentException("InputStream is null.");
    }

    final StringBuffer text = new StringBuffer();

    try {
      if (is.available() == 0) {
        return "";
      }
      try {
        OPCPackage container = OPCPackage.open(is);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(container);
        XSSFReader xssfReader = new XSSFReader(container);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        SheetTextExtractor sheetExtractor = new SheetTextExtractor(text);
        int parsedTabs = 0;
        while (iter.hasNext() && parsedTabs < MAX_TABS) {
          InputStream stream = iter.next();
          text.append('\n');
          text.append(iter.getSheetName());
          text.append('\n');
          processSheet(sheetExtractor, styles, strings, stream);
          stream.close();
          parsedTabs++;
        }
      } catch (InvalidFormatException e) {
        throw new DocumentReadException("The format of the document to read is invalid.", e);
      } catch (SAXException e) {
        throw new DocumentReadException("Problem during the document parsing.", e);
      } catch (OpenXML4JException e) {
        throw new DocumentReadException("Problem during the document parsing.", e);
      }

    } finally {
      if (is != null) {
        try {
          is.close();
        } catch (IOException e) {
          if (LOG.isTraceEnabled()) {
            LOG.trace("An exception occurred: " + e.getMessage());
          }
        }
      }
    }
    return text.toString();
  }

  public String getContentAsText(InputStream is, String encoding) throws IOException, DocumentReadException {
    // Ignore encoding
    return getContentAsText(is);
  }

  /*
   * (non-Javadoc)
   *
   * @see org.exoplatform.services.document.DocumentReader#getProperties(java.io.
   *      InputStream)
   */
  public Properties getProperties(final InputStream is) throws IOException, DocumentReadException {
    POIPropertiesReader reader = new POIPropertiesReader();
    reader.readDCProperties(SecurityHelper
                                .doPrivilegedIOExceptionAction(new PrivilegedExceptionAction<XSSFWorkbook>() {
                                  public XSSFWorkbook run() throws Exception {
                                    return new XSSFWorkbook(is);
                                  }
                                }));

    return reader.getProperties();
  }

  protected class SheetTextExtractor implements XSSFOptimizedSheetXMLHandler_03.SheetContentsHandler {
    private final StringBuffer output;
    private boolean firstCellOfRow = true;

    protected SheetTextExtractor(StringBuffer output) {
      this.output = output;
    }

    public void startRow(int rowNum) {
      firstCellOfRow = true;
    }

    public void endRow() {
      //output.append('\n');
    }

    public void cell(String cellRef, String formattedValue) {
      if (firstCellOfRow) {
        firstCellOfRow = false;
      } else {
        if (formattedValue != null) {
          output.append(' ');
        }
      }
      if (formattedValue != null) {
        output.append(formattedValue);
      }

    }

    public void headerFooter(String text, boolean isHeader, String tagName) {
      // We don't include headers in the output yet, so ignore
    }
  }
}