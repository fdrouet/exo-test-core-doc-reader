package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;


import org.apache.poi.hssf.eventusermodel.AbortableHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.*;
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
 * - after 5000 cells processed, we abort the parsing
 * <p/>
 * <p/>
 * we KEEP only the following data :
 * - tab name {@link org.apache.poi.hssf.record.BoundSheetRecord}
 * - cells with string (Strings which are not the result of a formula) ({@link org.apache.poi.hssf.record.LabelSSTRecord}}
 * <p/>
 * <p/>
 * we SKIP the following data :
 * - cells with number (date formatted or simple number) ({@link org.apache.poi.hssf.record.NumberRecord}}
 * - cells with blank value ({@link org.apache.poi.hssf.record.BlankRecord}}
 * - cells with boolean or error value ({@link org.apache.poi.hssf.record.BoolErrRecord}}
 * - cells with formula ({@link org.apache.poi.hssf.record.FormulaRecord}}
 */
public class MSExcelDocumentReader_Stream_03 extends MSExcelDocumentReader {
  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSExcelDocumentReader_Stream_02");
  private static final int MAX_CELL = 5000;

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

        public int cellnum = 0;

        // SSTRecords store a array of unique strings used in Excel.
        private SSTRecord sstrec;

        @Override
        public short abortableProcessRecord(Record record) {
          if (cellnum < MAX_CELL) {
            switch (record.getSid()) {
              // ## SKIP cells containing Numbers (Contains a numeric cell value.)
              case NumberRecord.sid:
                // NumberRecord numrec = (NumberRecord) record;
                cellnum++;
                break;
              // ## SKIP blank cells
              case BlankRecord.sid:
                // BlankRecord blankrec = (BlankRecord) record;
                break;
              // SKIP formula cells
              case FormulaRecord.sid:
                // FormulaRecord formrec = (FormulaRecord) record;
                cellnum++;
                break;
              case BoolErrRecord.sid:
                // BoolErrRecord boolrec = (BoolErrRecord) record;
                cellnum++;
                break;
              // SSTRecords store a array of unique strings used in Excel.
              case SSTRecord.sid:
                sstrec = (SSTRecord) record;
                break;
              case LabelSSTRecord.sid:
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                builder.append(sstrec.getString(lrec.getSSTIndex())).append(" ");
                cellnum++;
                break;
              case StringRecord.sid:
                // StringRecord sr = (StringRecord) record;
                cellnum++;
                break;
              case RowRecord.sid:
                // RowRecord rowrec = (RowRecord) record;
                break;
              case BoundSheetRecord.sid:
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                builder.append(bsr.getSheetname()).append(" ");
                break;
              case EOFRecord.sid:
                break;
            }
            // continue to process cells
            return 0;
          } else {
            LOG.info("#### " + cellnum + " indexed");
            // stop cells processing
            return -1;
          }
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