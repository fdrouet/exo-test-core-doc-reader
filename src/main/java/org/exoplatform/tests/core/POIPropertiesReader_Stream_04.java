package org.exoplatform.tests.core;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLPropertiesTextExtractor;
import org.apache.poi.hpsf.MarkUnsupportedException;
import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.PropertySetFactory;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.openxml4j.util.Nullable;
import org.apache.poi.poifs.eventfilesystem.POIFSReader;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;
import org.exoplatform.commons.utils.SecurityHelper;
import org.exoplatform.services.document.DCMetaData;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.impl.POIPropertiesReader;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;

import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.security.PrivilegedExceptionAction;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.TimeZone;

/**
 * Created by The eXo Platform SAS .
 * 
 * @author Gennady Azarenkov
 * @version $Id: $
 */

public class POIPropertiesReader_Stream_04 extends POIPropertiesReader
{

   private static final Log LOG = ExoLogger.getLogger("exo.core.component.document.POIPropertiesReader_Stream_04");

   private final Properties props = new Properties();

   public Properties getProperties()
   {
      return props;
   }

  public Properties readDCProperties(POIXMLDocument document) throws IOException, DocumentReadException
  {

//    POIXMLPropertiesTextExtractor extractor = new POIXMLPropertiesTextExtractor(document);
    POIXMLProperties extractor=document.getProperties();
    CoreProperties coreProperties = extractor.getCoreProperties();

    Nullable<String> lastModifiedBy = coreProperties.getUnderlyingProperties().getLastModifiedByProperty();

    SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
    df.setTimeZone(TimeZone.getDefault());

    if (lastModifiedBy != null && lastModifiedBy.getValue() != null && lastModifiedBy.getValue().length() > 0)
    {
      props.put(DCMetaData.CONTRIBUTOR, lastModifiedBy.getValue());
    }
    if (coreProperties.getDescription() != null && coreProperties.getDescription().length() > 0)
    {
      props.put(DCMetaData.DESCRIPTION, coreProperties.getDescription());
    }
    if (coreProperties.getCreated() != null)
    {
      try
      {
        Date d = df.parse(coreProperties.getUnderlyingProperties().getCreatedPropertyString());
        props.put(DCMetaData.DATE, d);
      }
      catch (ParseException e)
      {
        throw new DocumentReadException("Incorrect creation date: " + e.getMessage(), e);
      }
    }
    if (coreProperties.getCreator() != null && coreProperties.getCreator().length() > 0)
    {
      props.put(DCMetaData.CREATOR, coreProperties.getCreator());
    }
    if (coreProperties.getSubject() != null && coreProperties.getSubject().length() > 0)
    {
      props.put(DCMetaData.SUBJECT, coreProperties.getSubject());
    }
    if (coreProperties.getModified() != null)
    {
      try
      {
        Date d = df.parse(coreProperties.getUnderlyingProperties().getModifiedPropertyString());
        props.put(DCMetaData.DATE, d);
      }
      catch (ParseException e)
      {
        throw new DocumentReadException("Incorrect modification date: " + e.getMessage(), e);
      }
    }
    if (coreProperties.getSubject() != null && coreProperties.getSubject().length() > 0)
    {
      props.put(DCMetaData.SUBJECT, coreProperties.getSubject());
    }
    if (coreProperties.getTitle() != null && coreProperties.getTitle().length() > 0)
    {
      props.put(DCMetaData.TITLE, coreProperties.getTitle());
    }

    return props;
  }
   /**
    * Metadata extraction from ooxml documents (MS 2007 office file formats)
    * 
    * @param xmlProps
    * @return
    * @throws IOException
    * @throws DocumentReadException
    */
   public Properties readDCProperties(POIXMLProperties xmlProps) throws IOException, DocumentReadException
   {

      //POIXMLPropertiesTextExtractor extractor = new POIXMLPropertiesTextExtractor(document);

      CoreProperties coreProperties = xmlProps.getCoreProperties();

      Nullable<String> lastModifiedBy = coreProperties.getUnderlyingProperties().getLastModifiedByProperty();

      SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
      df.setTimeZone(TimeZone.getDefault());

      if (lastModifiedBy != null && lastModifiedBy.getValue() != null && lastModifiedBy.getValue().length() > 0)
      {
         props.put(DCMetaData.CONTRIBUTOR, lastModifiedBy.getValue());
      }
      if (coreProperties.getDescription() != null && coreProperties.getDescription().length() > 0)
      {
         props.put(DCMetaData.DESCRIPTION, coreProperties.getDescription());
      }
      if (coreProperties.getCreated() != null)
      {
         try
         {
            Date d = df.parse(coreProperties.getUnderlyingProperties().getCreatedPropertyString());
            props.put(DCMetaData.DATE, d);
         }
         catch (ParseException e)
         {
            throw new DocumentReadException("Incorrect creation date: " + e.getMessage(), e);
         }
      }
      if (coreProperties.getCreator() != null && coreProperties.getCreator().length() > 0)
      {
         props.put(DCMetaData.CREATOR, coreProperties.getCreator());
      }
      if (coreProperties.getSubject() != null && coreProperties.getSubject().length() > 0)
      {
         props.put(DCMetaData.SUBJECT, coreProperties.getSubject());
      }
      if (coreProperties.getModified() != null)
      {
         try
         {
            Date d = df.parse(coreProperties.getUnderlyingProperties().getModifiedPropertyString());
            props.put(DCMetaData.DATE, d);
         }
         catch (ParseException e)
         {
            throw new DocumentReadException("Incorrect modification date: " + e.getMessage(), e);
         }
      }
      if (coreProperties.getSubject() != null && coreProperties.getSubject().length() > 0)
      {
         props.put(DCMetaData.SUBJECT, coreProperties.getSubject());
      }
      if (coreProperties.getTitle() != null && coreProperties.getTitle().length() > 0)
      {
         props.put(DCMetaData.TITLE, coreProperties.getTitle());
      }

      return props;
   }

}
