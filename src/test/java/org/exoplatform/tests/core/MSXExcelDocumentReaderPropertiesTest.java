package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.util.Enumeration;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;


import com.carrotsearch.junitbenchmarks.BenchmarkOptions;
import com.carrotsearch.junitbenchmarks.BenchmarkRule;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.DocumentReader;
import org.exoplatform.services.document.impl.MSXExcelDocumentReader;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.FixMethodOrder;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestRule;
import org.junit.runners.MethodSorters;

/**
 * Test the performance of {@link org.exoplatform.services.document.impl.MSXExcelDocumentReader} with a new implementation
 */
//@AxisRange(min = 0, max = 1)
//@BenchmarkMethodChart(filePrefix = "benchmark-lists")
//@BenchmarkHistoryChart(filePrefix = "benchmark-history", maxRuns = 10, labelWith = LabelType.CUSTOM_KEY)
//@BenchmarkOptions(benchmarkRounds = 10, warmupRounds = 5, concurrency = -1, callgc = true)
@BenchmarkOptions(benchmarkRounds = 1, warmupRounds = 0, concurrency = -1, callgc = true)
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class MSXExcelDocumentReaderPropertiesTest {

  public static final String MS_XLSX_500KB = "MS-XLSX_500KB.xlsx";
  public static final String MS_XLSX_11MB_FORMULA = "MS_XLSX_11MB-formula.xlsx";
  public static final String MS_XLSX_18MB_FORMULA = "MS-XLSX_18MB-lot-of-formula.xlsx";
  public static final String MS_XLSX_METADATA = "test.xlsx";

  public static final String MS_XLSX_2_USE = MS_XLSX_18MB_FORMULA;

  public static final String TEST_LABEL = "test_" + MS_XLSX_2_USE;

  private final NumberFormat nf = NumberFormat.getInstance(Locale.FRENCH);

  @Rule
  public TestRule benchmarkRun = new BenchmarkRule();

  private DocumentReader docReaderORI;

  private DocumentReader docReaderStream04;

  private static final Map<String, Map<String, Map<String, Object>>> moreInfos = new TreeMap<String, Map<String, Map<String, Object>>>();

  @Before
  public void setUp() {
    docReaderORI = new MSXExcelDocumentReader();
    docReaderStream04 = new MSXExcelDocumentReader_StreamProperties_04();
    System.gc();
  }

  @Test
  public void test_XLS_ORI() throws IOException, DocumentReadException {
    final String version = "ORI";
    InputStream docIS = MSExcelDocumentReaderStreamTest.class.getResourceAsStream("/" + MS_XLSX_2_USE);
    long startUsedMemory = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory();
    Properties properties = docReaderORI.getProperties(docIS);
    docIS.close();
    addMoreInfos_memory(TEST_LABEL, version, startUsedMemory);
    addMoreInfos(TEST_LABEL, version, "properties size", properties.keySet().size() + "");
    addMoreInfos(TEST_LABEL, version, "content", properties);
  }

  @Test
  public void test_XLS_STREAM_04() throws IOException, DocumentReadException {
    final String version = "STREAM_04";
    InputStream docIS = MSExcelDocumentReaderStreamTest.class.getResourceAsStream("/" + MS_XLSX_2_USE);
    long startUsedMemory = Runtime.getRuntime().totalMemory()-Runtime.getRuntime().freeMemory();
    Properties properties = docReaderStream04.getProperties(docIS);
    docIS.close();
    addMoreInfos_memory(TEST_LABEL, version, startUsedMemory);
    addMoreInfos(TEST_LABEL, version, "properties size", properties.keySet().size() + "");
    addMoreInfos(TEST_LABEL, version, "content", properties);
  }

  private void addMoreInfos_memory(String test, String version, long startUsedMemory) {
    long mem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory() - startUsedMemory;
    addMoreInfos(test, version, "used memory", nf.format(mem) + " (" + mem + ")");
  }

  private void addMoreInfos(String test, String version, String type, Object value) {
    StringBuffer buffer = new StringBuffer();

    if (value instanceof Properties) {
      Properties props = (Properties)value;
      Enumeration<Object> keys = props.keys();
      while(keys.hasMoreElements()){
        Object key = keys.nextElement();
        buffer.append(key);
        buffer.append(" : ");
        buffer.append(((Properties)value).get(key));
        buffer.append("\n");
      }
    } else {
      buffer.append(value.toString());
    }

    addMoreInfos(test, version, type, buffer.toString());
  }
  private void addMoreInfos(String test, String version, String type, String value) {
    Map<String, Object> versionData;
    if (moreInfos.containsKey(test)) {
      Map<String, Map<String, Object>> testData = moreInfos.get(test);

      if (testData.containsKey(version)) {
        versionData = testData.get(version);
      } else {
        versionData = new TreeMap<String, Object>();
        testData.put(version, versionData);
      }
    } else {
      Map<String, Map<String, Object>> testData = new TreeMap<String, Map<String, Object>>();
      moreInfos.put(test, testData);
      versionData = new TreeMap<String, Object>();
      testData.put(version, versionData);
    }

    versionData.put(type, value);
  }

  @After
  public void tearDown() {
    System.gc();
  }

  @AfterClass
  public static void afterTests() {
    System.out.println("## MORE INFOS:\n");
    for (Map.Entry<String, Map<String, Map<String, Object>>> entryTest : moreInfos.entrySet()) {
      String testName = entryTest.getKey();
      System.out.println("##    TEST:" + testName);
      Map<String, Map<String, Object>> testData = entryTest.getValue();
      for (Map.Entry<String, Map<String, Object>> entryVersion : testData.entrySet()) {
        String testVersion = entryVersion.getKey();
        Map<String, Object> testInfos = entryVersion.getValue();
        System.out.println("##       " + testVersion + ": properties size = " + testInfos.get("properties size"));
        System.out.println("##       " + testVersion + ": used memory  = " + testInfos.get("used memory"));
        System.out.println("##       " + testVersion + ": content    = \n=====\n" + testInfos.get("content") + "\n=====");
      }
    }
  }
}
