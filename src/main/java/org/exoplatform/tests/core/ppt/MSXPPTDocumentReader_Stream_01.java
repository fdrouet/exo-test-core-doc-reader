package org.exoplatform.tests.core.ppt;

import java.io.IOException;
import java.io.InputStream;
import java.security.PrivilegedAction;
import java.security.PrivilegedActionException;
import java.security.PrivilegedExceptionAction;


import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.xmlbeans.XmlException;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.MSXPPTDocumentReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

/**
 * Patched original eXo MS Powerpoint Document Reader with the following changes .
 */
public class MSXPPTDocumentReader_Stream_01 extends MSXPPTDocumentReader {

  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSXPPTDocumentReader_Stream_01");


  /**
   * Returns only a text from .pptx file content.
   *
   * @param is an input stream with .pptx file content.
   * @return The string only with text from file content.
   */
  public String getContentAsText(final InputStream is) throws IOException, DocumentReadException
  {
    if (is == null)
    {
      throw new IllegalArgumentException("InputStream is null.");
    }
    try
    {
      if (is.available() == 0)
      {
        return "";
      }

      final XSLFPowerPointExtractor ppe;
      try
      {
        ppe = SecurityHelper.doPrivilegedExceptionAction(new PrivilegedExceptionAction<XSLFPowerPointExtractor>() {
          public XSLFPowerPointExtractor run() throws Exception {
            return new XSLFPowerPointExtractor(OPCPackage.open(is));
          }
        });
      }
      catch (PrivilegedActionException pae)
      {
        Throwable cause = pae.getCause();
        if (cause instanceof IOException)
        {
          throw new DocumentReadException("Can't open presentation.", cause);
        }
        else if (cause instanceof OpenXML4JRuntimeException)
        {
          throw new DocumentReadException("Can't open presentation.", cause);
        }
        else if (cause instanceof OpenXML4JException)
        {
          throw new DocumentReadException("Can't open presentation.", cause);
        }
        else if (cause instanceof XmlException)
        {
          throw new DocumentReadException("Can't open presentation.", cause);
        }
        else if (cause instanceof RuntimeException)
        {
          throw (RuntimeException)cause;
        }
        else
        {
          throw new RuntimeException(cause);
        }
      }
      return SecurityHelper.doPrivilegedAction(new PrivilegedAction<String>() {
        public String run() {
          return ppe.getText(true, true);
        }
      });
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
  }


}