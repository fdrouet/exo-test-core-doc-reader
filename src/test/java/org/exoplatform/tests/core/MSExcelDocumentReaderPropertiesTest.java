package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;


import com.carrotsearch.junitbenchmarks.BenchmarkOptions;
import com.carrotsearch.junitbenchmarks.BenchmarkRule;
import org.exoplatform.services.document.DocumentReadException;
import org.exoplatform.services.document.DocumentReader;
import org.exoplatform.services.document.impl.MSExcelDocumentReader;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.FixMethodOrder;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestRule;
import org.junit.runners.MethodSorters;

/**
 * Test the performance of {@link org.exoplatform.services.document.impl.MSExcelDocumentReader#getContentAsText(java.io.InputStream)} with a new implementation
 */
//@AxisRange(min = 0, max = 1)
//@BenchmarkMethodChart(filePrefix = "benchmark-lists")
//@BenchmarkHistoryChart(filePrefix = "benchmark-history", maxRuns = 10, labelWith = LabelType.CUSTOM_KEY)
@BenchmarkOptions(benchmarkRounds = 1, warmupRounds = 0, concurrency = -1, callgc = true)
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class MSExcelDocumentReaderPropertiesTest {

  public static final String MS_XLS_25KB_LIGHT = "MS-XLS_25KB.xls";
  public static final String MS_XLS_500KB = "MS-XLS_500KB.xls";
  public static final String MS_XLS_TEXT = "MS-XLS_17MB-lot-of-text.xls";
  public static final String MS_XLS_FORMULA = "MS-XLS_27MB-lot-of-formula.xls";

  public static final String MS_XLS_2_USE = MS_XLS_FORMULA;

  public static final String TEST_LABEL = "test_"+MS_XLS_2_USE;

  private final NumberFormat nf = NumberFormat.getInstance(Locale.FRENCH);

  @Rule
  public TestRule benchmarkRun = new BenchmarkRule();

  private DocumentReader docReaderORI;

  private DocumentReader docReaderPatched01;

  private static final Map<String, Map<String, Map<String, Object>>> moreInfos = new TreeMap<String, Map<String, Map<String, Object>>>();

  @Before
  public void setUp() {
    docReaderORI = new MSExcelDocumentReader();
    docReaderPatched01 = new MSExcelDocumentReader_Patched_01();
    System.gc();
  }

  @Test
  public void test_XLS_ORI() throws IOException, DocumentReadException {
    final String version = "ORI";
    InputStream docIS = MSExcelDocumentReaderPropertiesTest.class.getResourceAsStream("/" + MS_XLS_2_USE);
    long startUsedMemory = Runtime.getRuntime().totalMemory()-Runtime.getRuntime().freeMemory();
    Properties properties = docReaderORI.getProperties(docIS);
    docIS.close();
    addMoreInfos_memory(TEST_LABEL, version, startUsedMemory);
    addMoreInfos(TEST_LABEL, version, "properties size", properties.keySet().size() + "");
    addMoreInfos(TEST_LABEL, version, "content", properties);
  }

  @Test
  public void test_XLS_PATCHED_01() throws IOException, DocumentReadException {
    final String version = "PATCHED_01";
    InputStream docIS = MSExcelDocumentReaderPropertiesTest.class.getResourceAsStream("/" + MS_XLS_2_USE);
    long startUsedMemory = Runtime.getRuntime().totalMemory()-Runtime.getRuntime().freeMemory();
    Properties properties = docReaderPatched01.getProperties(docIS);
    docIS.close();
    addMoreInfos_memory(TEST_LABEL, version, startUsedMemory);
    addMoreInfos(TEST_LABEL, version, "properties size", properties.keySet().size() + "");
    addMoreInfos(TEST_LABEL, version, "content", properties);
  }

  private void addMoreInfos_memory (String test, String version, long startUsedMemory) {
    long mem = Runtime.getRuntime().totalMemory() - Runtime.getRuntime().freeMemory() - startUsedMemory;
    addMoreInfos(test, version, "used memory", nf.format(mem)+" ("+mem+")");
  }

  private void addMoreInfos(String test, String version, String type, Object value) {
    StringBuffer buffer = new StringBuffer();

    if (value instanceof Properties) {
      for (String key: ((Properties)value).stringPropertyNames()){
        buffer.append(key);
        buffer.append(" : ");
        buffer.append(((Properties)value).getProperty(key));
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
