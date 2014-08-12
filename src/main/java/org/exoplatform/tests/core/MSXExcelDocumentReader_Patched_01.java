package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.security.PrivilegedExceptionAction;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.BaseDocumentReader;
import org.exoplatform.services.document.impl.POIPropertiesReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

/**
 * Patched original eXo MS Excel Document Reader with the following changes .
 *
 * we only index a maximum of 5000 cells:
 * - at most 1000 cells per sheet
 * - at most the 5 first tabs of the spreadsheet
 *
 * we KEEP only the following data :
 * - cells with number (date formatted or simple number)
 * - cells with string
 *
 * we SKIP the following data :
 * - cells with blank value
 * - cells with boolean value
 * - cells with formula
 * - cells with error
 *
 */
public class MSXExcelDocumentReader_Patched_01 extends BaseDocumentReader
{

  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSXExcelDocumentReader");

  private static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss.SSSZ";

  private static final int MAX_TAB = 5;

  private static final int MAX_CELL = 1000;

  /**
   * @see org.exoplatform.services.document.DocumentReader#getMimeTypes()
   */
  public String[] getMimeTypes()
  {
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

  /**
   * Returns only a text from .xlsx file content.
   *
   * @param is an input stream with .xls file content.
   * @return The string only with text from file content.
   */
  public String getContentAsText(final InputStream is) throws IOException, DocumentReadException
  {
    if (is == null)
    {
      throw new IllegalArgumentException("InputStream is null.");
    }

    StringBuilder builder = new StringBuilder("");
    SimpleDateFormat dateFormat = new SimpleDateFormat(DATE_FORMAT);

    try
    {
      if (is.available() == 0)
      {
        return "";
      }

      XSSFWorkbook wb;
      try
      {
        wb = SecurityHelper.doPrivilegedIOExceptionAction(new PrivilegedExceptionAction<XSSFWorkbook>()
        {
          public XSSFWorkbook run() throws Exception
          {
            return new XSSFWorkbook(is);
          }
        });
      }
      catch (IOException e)
      {
        throw new DocumentReadException("Can't open spreadsheet.", e);
      }
      catch (OpenXML4JRuntimeException e)
      {
        return builder.toString();
      }
      for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets() && sheetNum < MAX_TAB ; sheetNum++)
      {
        XSSFSheet sheet = wb.getSheetAt(sheetNum);
        if (sheet != null)
        {
          int countCell = MAX_CELL;
          for (int rowNum = sheet.getFirstRowNum(); rowNum <= sheet.getLastRowNum() && countCell > 0 ; rowNum++)
          {
            XSSFRow row = sheet.getRow(rowNum);

            if (row != null)
            {
              int lastcell = row.getLastCellNum();
              for (int k = 0; k < lastcell && countCell > 0; k++)
              {
                XSSFCell cell = row.getCell(k);
                countCell --;
                if (cell != null)
                {
                  switch (cell.getCellType())
                  {
                    case XSSFCell.CELL_TYPE_NUMERIC : {
                      double d = cell.getNumericCellValue();
                      if (isCellDateFormatted(cell))
                      {
                        Date date = HSSFDateUtil.getJavaDate(d);
                        String cellText = dateFormat.format(date);
                        builder.append(cellText).append(" ");
                      }
                      else
                      {
                        builder.append(d).append(" ");
                      }
                      break;
                    }
                    case XSSFCell.CELL_TYPE_STRING :
                      builder.append(cell.getStringCellValue().toString()).append(" ");
                      break;
                    default :
                      break;
                  }
                }
              }
            }
          }
        }
      }
    }
    finally
    {
      if (is != null)
      {
        try
        {
          is.close();
        }
        catch (IOException e)
        {
          if (LOG.isTraceEnabled())
          {
            LOG.trace("An exception occurred: " + e.getMessage());
          }
        }
      }
    }
    return builder.toString();
  }

  public String getContentAsText(InputStream is, String encoding) throws IOException, DocumentReadException
  {
    // Ignore encoding
    return getContentAsText(is);
  }

  /*
   * (non-Javadoc)
   *
   * @see org.exoplatform.services.document.DocumentReader#getProperties(java.io.
   *      InputStream)
   */
  public Properties getProperties(final InputStream is) throws IOException, DocumentReadException
  {
    POIPropertiesReader reader = new POIPropertiesReader();
    reader.readDCProperties(SecurityHelper
                                .doPrivilegedIOExceptionAction(new PrivilegedExceptionAction<XSSFWorkbook>()
                                {
                                  public XSSFWorkbook run() throws Exception
                                  {
                                    return new XSSFWorkbook(is);
                                  }
                                }));

    return reader.getProperties();
  }

  public static boolean isCellDateFormatted(XSSFCell cell)
  {
    boolean bDate = false;
    double d = cell.getNumericCellValue();
    if (HSSFDateUtil.isValidExcelDate(d))
    {
      XSSFCellStyle style = cell.getCellStyle();
      int i = style.getDataFormat();
      switch (i)
      {
        case 0xe : // m/d/yy
        case 0xf : // d-mmm-yy
        case 0x10 : // d-mmm
        case 0x11 : // mmm-yy
        case 0x12 : // h:mm AM/PM
        case 0x13 : // h:mm:ss AM/PM
        case 0x14 : // h:mm
        case 0x15 : // h:mm:ss
        case 0x16 : // m/d/yy h:mm
        case 0x2d : // mm:ss
        case 0x2e : // [h]:mm:ss
        case 0x2f : // mm:ss.0

        case 0xa5 : // ??
        case 0xa7 : // ??
        case 0xa9 : // ??

        case 0xac : // mm:dd:yy not specified in javadoc
        case 0xad : // yyyy-mm-dd not specified in javadoc
        case 0xae : // mm:dd:yyyy not specified in javadoc
        case 0xaf : // m:d:yy not specified in javadoc
          bDate = true;
          break;
        default :
          bDate = false;
          break;
      }
    }
    return bDate;
  }
}