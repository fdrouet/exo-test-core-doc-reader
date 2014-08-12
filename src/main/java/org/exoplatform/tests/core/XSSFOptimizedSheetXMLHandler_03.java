package org.exoplatform.tests.core;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * This class handles the processing of a sheet#.xml
 * sheet part of a XSSF .xlsx file, and generates
 * row and cell events for it.
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
public class XSSFOptimizedSheetXMLHandler_03 extends DefaultHandler {

  private static final Log LOG = ExoLogger.getExoLogger("exo.core.component.document.XSSFOptimizedSheetXMLHandler_03");

  /**
   * These are the different kinds of cells we support.
   * We keep track of the current one between
   * the start and end.
   */
  enum xssfDataType {
    BOOLEAN,
    ERROR,
    FORMULA,
    INLINE_STRING,
    SST_STRING,
    NUMBER,
  }

  /**
   * Table with the styles used for formatting
   */
  private StylesTable stylesTable;

  private ReadOnlySharedStringsTable sharedStringsTable;

  /**
   * The maximum number of cells to parse in the Sheet (-1 mean All cells in the sheet)
   */
  private long maxCellsToParse = -1;
  private long currentCellsParsed = 0;

  /**
   * Where our text is going
   */
  private final SheetContentsHandler output;

  // Set when V start element is seen
  private boolean vIsOpen;
  // Set when F start element is seen
  private boolean fIsOpen;
  // Set when an Inline String "is" is seen
  private boolean isIsOpen;
  // Set when a header/footer element is seen
  private boolean hfIsOpen;

  // Set when cell start element is seen;
  // used when cell close element is seen.
  private xssfDataType nextDataType;

  // Used to format numeric cell values.
  private short formatIndex;
  private String formatString;
  private final DataFormatter formatter = new DataFormatter();
  private String cellRef;

  // Gathers characters as they are seen.
  private StringBuffer value = new StringBuffer();
  private StringBuffer formula = new StringBuffer();
  private StringBuffer headerFooter = new StringBuffer();

  /**
   * Accepts objects needed while parsing.
   *
   * @param styles  Table of styles
   * @param strings Table of shared strings
   */
  public XSSFOptimizedSheetXMLHandler_03(
      StylesTable styles,
      ReadOnlySharedStringsTable strings,
      SheetContentsHandler sheetContentsHandler,
      long maxCellsToParse) {
    this.stylesTable = styles;
    this.sharedStringsTable = strings;
    this.output = sheetContentsHandler;
    this.nextDataType = xssfDataType.NUMBER;
    this.maxCellsToParse = maxCellsToParse;
  }

  private boolean isTextTag(String name) {
    if ("v".equals(name)) {
      // Easy, normal v text tag
      return true;
    }
    if ("inlineStr".equals(name)) {
      // Easy inline string
      return true;
    }
    if ("t".equals(name) && isIsOpen) {
      // Inline string <is><t>...</t></is> pair
      return true;
    }
    // It isn't a text tag
    return false;
  }

  public void startElement(String uri, String localName, String name,
                           Attributes attributes) throws SAXException {

    if (isTextTag(name)) {
      vIsOpen = true;
      // Clear contents cache
      value.setLength(0);
    } else if ("is".equals(name)) {
      // Inline string outer tag
      isIsOpen = true;
    } else if ("f".equals(name)) {
      // Clear contents cache
      formula.setLength(0);

      // Mark us as being a formula if not already
      if (nextDataType == xssfDataType.NUMBER) {
        nextDataType = xssfDataType.FORMULA;
      }

      // Decide where to get the formula string from
      String type = attributes.getValue("t");
      if (type != null && type.equals("shared")) {
        // Is it the one that defines the shared, or uses it?
        String ref = attributes.getValue("ref");
//        String si = attributes.getValue("si");

        if (ref != null) {
          // This one defines it
          // TODO Save it somewhere
          fIsOpen = true;
        } else {
          // This one uses a shared formula
          // TODO Retrieve the shared formula and tweak it to match the current cell
          //System.err.println("Warning - shared formulas not yet supported!");
        }
      } else {
        fIsOpen = true;
      }
    } else if ("oddHeader".equals(name) || "evenHeader".equals(name) ||
        "firstHeader".equals(name) || "firstFooter".equals(name) ||
        "oddFooter".equals(name) || "evenFooter".equals(name)) {
      hfIsOpen = true;
      // Clear contents cache
      headerFooter.setLength(0);
    } else if ("row".equals(name)) {
      int rowNum = Integer.parseInt(attributes.getValue("r")) - 1;
      output.startRow(rowNum);
    }
    // c => cell
    else if ("c".equals(name)) {
      // Set up defaults.
      this.nextDataType = xssfDataType.NUMBER;
      this.formatIndex = -1;
      this.formatString = null;
      cellRef = attributes.getValue("r");
      String cellType = attributes.getValue("t");
      String cellStyleStr = attributes.getValue("s");
      if ("b".equals(cellType))
        nextDataType = xssfDataType.BOOLEAN;
      else if ("e".equals(cellType))
        nextDataType = xssfDataType.ERROR;
      else if ("inlineStr".equals(cellType))
        nextDataType = xssfDataType.INLINE_STRING;
      else if ("s".equals(cellType))
        nextDataType = xssfDataType.SST_STRING;
      else if ("str".equals(cellType))
        nextDataType = xssfDataType.FORMULA;
      else if (cellStyleStr != null) {
        // Number, but almost certainly with a special style or format
        int styleIndex = Integer.parseInt(cellStyleStr);
        XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
        this.formatIndex = style.getDataFormat();
        this.formatString = style.getDataFormatString();
        if (this.formatString == null)
          this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
      }
    }
  }

  public void endElement(String uri, String localName, String name)
      throws SAXException {
    String thisStr = null;

    // v => contents of a cell
    if (isTextTag(name)) {
      vIsOpen = false;

      // Process the value contents as required, now we have it all
      switch (nextDataType) {
        case BOOLEAN:
          currentCellsParsed++;
          break;
        case ERROR:
          currentCellsParsed++;
          break;

        case FORMULA:
          currentCellsParsed++;
          break;

        case INLINE_STRING:
          // TODO: Can these ever have formatting on them?
          XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
          thisStr = rtsi.toString();
          currentCellsParsed++;
          break;

        case SST_STRING:
          String sstIndex = value.toString();
          try {
            int idx = Integer.parseInt(sstIndex);
            XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
            thisStr = rtss.toString();
          } catch (NumberFormatException ex) {
            System.err.println("Failed to parse SST index '" + sstIndex + "': " + ex.toString());
          }
          currentCellsParsed++;
          break;

        case NUMBER:
          currentCellsParsed++;
          break;

        default:
          thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
          currentCellsParsed++;
          break;
      }

      // Output
      output.cell(cellRef, thisStr);
    } else if ("f".equals(name)) {
      fIsOpen = false;
    } else if ("is".equals(name)) {
      isIsOpen = false;
    } else if ("row".equals(name)) {
      output.endRow();
    } else if ("oddHeader".equals(name) || "evenHeader".equals(name) ||
        "firstHeader".equals(name)) {
      hfIsOpen = false;
      output.headerFooter(headerFooter.toString(), true, name);
    } else if ("oddFooter".equals(name) || "evenFooter".equals(name) ||
        "firstFooter".equals(name)) {
      hfIsOpen = false;
      output.headerFooter(headerFooter.toString(), false, name);
    }
    if (maxCellsToParse >= 0 && currentCellsParsed > maxCellsToParse) {
      LOG.info(this.toString() + " - We stop parsing the sheet after " + (currentCellsParsed - 1) + " parsed cells");
      throw new StopSheetParsingException("Maximum number of cells to parse reached");
    }
  }

  /**
   * Captures characters only if a suitable element is open.
   * Originally was just "v"; extended for inlineStr also.
   */
  public void characters(char[] ch, int start, int length)
      throws SAXException {
    if (vIsOpen) {
      value.append(ch, start, length);
    }
    if (fIsOpen) {
      formula.append(ch, start, length);
    }
    if (hfIsOpen) {
      headerFooter.append(ch, start, length);
    }
  }

  /**
   * You need to implement this to handle the results
   * of the sheet parsing.
   */
  public interface SheetContentsHandler {
    /**
     * A row with the (zero based) row number has started
     */
    public void startRow(int rowNum);

    /**
     * A row with the (zero based) row number has ended
     */
    public void endRow();

    /**
     * A cell, with the given formatted value, was encountered
     */
    public void cell(String cellReference, String formattedValue);

    /**
     * A header or footer has been encountered
     */
    public void headerFooter(String text, boolean isHeader, String tagName);
  }

  /**
   * This exception is used to ask the underlying SAX Parser to stop parsing a XML sheet.
   */
  public class StopSheetParsingException extends SAXException {
    public StopSheetParsingException() {
    }

    public StopSheetParsingException(String message) {
      super(message);
    }

    public StopSheetParsingException(Exception e) {
      super(e);
    }

    public StopSheetParsingException(String message, Exception e) {
      super(message, e);
    }
  }
}
