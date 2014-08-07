package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;


import org.apache.poi.hssf.eventusermodel.AbortableHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.common.UnicodeString;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.MSExcelDocumentReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

/**
 * Stream based MS Excel Document Reader.
 * <p/>
 * we only index :
 * - a maximum of 5000 cells
 * - a maximum of 1000 cells per tab
 * - a maximum of 5 tabs
 * <p/>
 * <p/>
 * we KEEP only the following data :
 * - tab name {@link org.apache.poi.hssf.record.BoundSheetRecord}
 * - cells with string with a length > 2 chars (Strings which are not the result of a formula) ({@link org.apache.poi.hssf.record.LabelSSTRecord}}
 * <p/>
 * <p/>
 * we SKIP the following data :
 * - cells with number (date formatted or simple number) ({@link org.apache.poi.hssf.record.NumberRecord}}
 * - cells with blank value ({@link org.apache.poi.hssf.record.BlankRecord}}
 * - cells with boolean or error value ({@link org.apache.poi.hssf.record.BoolErrRecord}}
 * - cells with formula ({@link org.apache.poi.hssf.record.FormulaRecord}}
 */
public class MSExcelDocumentReader_Stream_05 extends MSExcelDocumentReader {
  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSExcelDocumentReader_Stream_05");
  private static final int MAX_TAB = 5;
  private static final int MAX_CELLTAB = 1000;

  @Override
  public String getContentAsText(InputStream is) throws IOException, DocumentReadException {
    if (is == null) {
      throw new IllegalArgumentException("InputStream is null.");
    }

    final StringBuilder builder = new StringBuilder("");

    try {
      if (is.available() == 0) {
        return "";
      }

      // lazy listen for ALL records with the listener shown above
      HSSFListener listener = new AbortableHSSFListener() {

        public int tabnum = 0;
        public int cellread = 0;
        public int celltab = 0;

        // SSTRecords store a array of unique strings used in Excel.
        private SSTRecord sstrec;

        @Override
        public short abortableProcessRecord(Record record) {
          if (tabnum > MAX_TAB) {
            // stop cells processing
            LOG.info("#### " + cellread + " indexed");
            return -1;
          }

          switch (record.getSid()) {
            // ## SKIP cells containing Numbers (Contains a numeric cell value.)
            case NumberRecord.sid:
              // NumberRecord numrec = (NumberRecord) record;
              if (celltab < MAX_CELLTAB) {
                celltab++;
              }
              cellread++;
              break;
            // ## SKIP blank cells
            case BlankRecord.sid:
              // BlankRecord blankrec = (BlankRecord) record;
              break;
            // SKIP formula cells
            case FormulaRecord.sid:
              // FormulaRecord formrec = (FormulaRecord) record;
              if (celltab < MAX_CELLTAB) {
                celltab++;
              }
              cellread++;
              break;
            case BoolErrRecord.sid:
              // BoolErrRecord boolrec = (BoolErrRecord) record;
              if (celltab < MAX_CELLTAB) {
                celltab++;
              }
              cellread++;
              break;
            // SSTRecords store a array of unique strings used in Excel.
            case SSTRecord.sid:
              sstrec = (SSTRecord) record;
              break;
            case LabelSSTRecord.sid:
              if (celltab < MAX_CELLTAB) {
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                UnicodeString lrecValue = sstrec.getString(lrec.getSSTIndex());
                if (lrecValue.getCharCount() > 2) {
                  builder.append(lrecValue).append(" ");
                }
                celltab++;
              }
              cellread++;
              break;
            case StringRecord.sid:
              // StringRecord sr = (StringRecord) record;
              if (celltab < MAX_CELLTAB) {
                celltab++;
              }
              cellread++;
              break;
            // the BOFRecord can represent either the beginning of a sheet or the workbook
            case BOFRecord.sid:
              BOFRecord bof = (BOFRecord) record;
              if (bof.getType() == bof.TYPE_WORKSHEET) {
                tabnum++;
                celltab = 0;
              }
              break;
            case BoundSheetRecord.sid:
              BoundSheetRecord bsr = (BoundSheetRecord) record;
              builder.append(bsr.getSheetname()).append(" ");
              break;
            case EOFRecord.sid:
              LOG.info("#### " + cellread + " indexed");
              break;
          }
          return 0;
        }
      };
      // create a new org.apache.poi.poifs.filesystem.Filesystem
      POIFSFileSystem poifs = new POIFSFileSystem(is);
      // get the Workbook (excel part) stream in a InputStream
      InputStream din = poifs.createDocumentInputStream("Workbook");
      // construct out HSSFRequest object
      HSSFRequest req = new HSSFRequest();
      req.addListenerForAllRecords(listener);
      // create our event factory
      HSSFEventFactory factory = new HSSFEventFactory();
      // process our events based on the document input stream
      factory.processEvents(req, din);
      // and our document input stream (don't want to leak these!)
      din.close();

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
    return builder.toString();
  }
}