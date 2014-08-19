package org.exoplatform.tests.core.ppt;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;


import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.POIPropertiesReader;
import org.exoplatform.services.document.impl.PPTDocumentReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

/**
 * Patched original eXo MS Powerpoint Document Reader with the following changes .
 */
public class MSPPTDocumentReader_Stream_01 extends PPTDocumentReader {

  private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.MSPPTDocumentReader_Stream_01");


  /**
   * Returns only a text from .ppt file content.
   *
   * @param is an input stream with .ppt file content.
   * @return The string only with text from file content.
   */
  public String getContentAsText(InputStream is) throws IOException, DocumentReadException
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

      PowerPointExtractor ppe;
      try
      {
        ppe = new PowerPointExtractor(is);
      }
      catch (IOException e)
      {
        throw new DocumentReadException("Can't open presentation.", e);
      }
      return ppe.getText(true, true);
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
  public Properties getProperties(InputStream is) throws IOException, DocumentReadException
  {
    POIPropertiesReader reader = new POIPropertiesReader();
    reader.readDCProperties(is);
    return reader.getProperties();
  }

}