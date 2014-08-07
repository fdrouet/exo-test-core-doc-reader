package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.security.PrivilegedAction;
import java.text.SimpleDateFormat;
import java.util.Date;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.MSExcelDocumentReader;
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
 * - cells with string ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING}}
 *
 * we SKIP the following data :
 * - cells with number (date formatted or simple number) ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC}}
 * - cells with blank value ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK}}
 * - cells with boolean value ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN}}
 * - cells with formula ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_FORMULA}}
 * - cells with error ({@link org.apache.poi.ss.usermodel.Cell.CELL_TYPE_ERROR}}
 *
 */
public class MSExcelDocumentReader_Patched_02 extends MSExcelDocumentReader {
  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSExcelDocumentReader_Patched");
  private static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss.SSSZ";
  private static final int MAX_TAB = 5;
  private static final int MAX_CELL = 1000;

  @Override
  public String getContentAsText(InputStream is) throws IOException, DocumentReadException {
    if (is == null)
    {
      throw new IllegalArgumentException("InputStream is null.");
    }

    final StringBuilder builder = new StringBuilder("");

    SimpleDateFormat dateFormat = new SimpleDateFormat(DATE_FORMAT);

    try
    {
      if (is.available() == 0)
      {
        return "";
      }

      HSSFWorkbook wb;
      try
      {
        wb = new HSSFWorkbook(is);
      }
      catch (IOException e)
      {
        throw new DocumentReadException("Can't open spreadsheet.", e);
      }
      for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets() && sheetNum < MAX_TAB; sheetNum++)
      {
        HSSFSheet sheet = wb.getSheetAt(sheetNum);
        if (sheet != null)
        {
          int countCell = MAX_CELL;
          for (int rowNum = sheet.getFirstRowNum(); rowNum <= sheet.getLastRowNum() && countCell > 0; rowNum++)
          {
            HSSFRow row = sheet.getRow(rowNum);

            if (row != null)
            {
              int lastcell = row.getLastCellNum();
              for (int k = 0; k < lastcell && countCell > 0; k++)
              {
                final HSSFCell cell = row.getCell((short)k);
                countCell --;
                if (cell != null)
                {
                  switch (cell.getCellType())
                  {
                    case HSSFCell.CELL_TYPE_STRING :
                      SecurityHelper.doPrivilegedAction(new PrivilegedAction<Void>() {
                        public Void run() {
                          builder.append(cell.getStringCellValue().toString()).append(" ");
                          return null;
                        }
                      });
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
}
