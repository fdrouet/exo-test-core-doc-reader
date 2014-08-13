package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.security.PrivilegedExceptionAction;
import java.util.Properties;


import org.apache.poi.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.POIPropertiesReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

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
public class MSXExcelDocumentReader_StreamProperties_04 extends MSXExcelDocumentReader_Stream_04 {

  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSXExcelDocumentReader_StreamProperties_04");
  /*
   * (non-Javadoc)
   *
   * @see org.exoplatform.services.document.DocumentReader#getProperties(java.io.
   *      InputStream)
   */
  public Properties getProperties(final InputStream is) throws IOException, DocumentReadException {

//    POIPropertiesReader reader = new POIPropertiesReader();
//    reader.readDCProperties(SecurityHelper
//                                .doPrivilegedIOExceptionAction(new PrivilegedExceptionAction<XSSFWorkbook>() {
//                                  public XSSFWorkbook run() throws Exception {
//                                    return new XSSFWorkbook(is);
//                                  }
//                                }));
//
//    return reader.getProperties();



    try {
      OPCPackage container = SecurityHelper
          .doPrivilegedIOExceptionAction(new PrivilegedExceptionAction<OPCPackage>() {
            public OPCPackage run() throws Exception {
              return OPCPackage.open(is);
            }
          });
      POIXMLProperties xmlProperties = new POIXMLProperties(container);
      POIPropertiesReader_Stream_04 reader = new POIPropertiesReader_Stream_04();
      reader.readDCProperties(xmlProperties);
      return reader.getProperties();
    } catch (InvalidFormatException e) {
      throw new DocumentReadException("The format of the document to read is invalid.", e);
    } catch (XmlException e) {
      throw new DocumentReadException("Problem during the document parsing.", e);
    } catch (OpenXML4JException e) {
      throw new DocumentReadException("Problem during the document parsing.", e);
    }
  }

}