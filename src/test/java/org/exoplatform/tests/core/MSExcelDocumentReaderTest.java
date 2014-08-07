package org.exoplatform.tests.core;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import java.util.TreeMap;


import com.carrotsearch.junitbenchmarks.BenchmarkOptions;
import com.carrotsearch.junitbenchmarks.BenchmarkRule;
import com.carrotsearch.junitbenchmarks.annotation.AxisRange;
import com.carrotsearch.junitbenchmarks.annotation.BenchmarkHistoryChart;
import com.carrotsearch.junitbenchmarks.annotation.BenchmarkMethodChart;
import com.carrotsearch.junitbenchmarks.annotation.LabelType;
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
 * Test the performance of {@link org.exoplatform.services.document.impl.MSExcelDocumentReader} with a new implementation
 */
@AxisRange(min = 0, max = 1)
@BenchmarkMethodChart(filePrefix = "benchmark-lists")
@BenchmarkHistoryChart(filePrefix = "benchmark-history", maxRuns = 10, labelWith = LabelType.CUSTOM_KEY)
@BenchmarkOptions(benchmarkRounds = 10, warmupRounds = 10, concurrency = -1, callgc = true)
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class MSExcelDocumentReaderTest {

  public static final String MS_XLS_25KB_LIGHT = "MS-XLS_25KB.xls";
  public static final String MS_XLS_500KB = "MS-XLS_500KB.xls";
  public static final String MS_XLS_TEXT = "MS-XLS_17MB-lot-of-text.xls";
  public static final String MS_XLS_FORMULA = "MS-XLS_27MB-lot-of-formula.xls";

  @Rule
  public TestRule benchmarkRun = new BenchmarkRule();

  private DocumentReader docReaderOld;

  private DocumentReader docReaderNew;

  private static final Map<String, Map<String, Map<String, Object>>> moreInfos = new TreeMap<String, Map<String, Map<String, Object>>>();

  @Before
  public void setUp() {
    docReaderOld = new MSExcelDocumentReader();
    docReaderNew = new MSExcelDocumentReader_Patched_01();
    System.gc();
  }

  @Test
  public void test_25KB_Light_ORI() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_25KB_LIGHT);
    String content = docReaderOld.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_25KB_Light", "ORI", "content size",content.length()+"");
  }

  @Test
  public void test_25KB_Light_PATCHED() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_25KB_LIGHT);
    String content = docReaderNew.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_25KB_Light", "PATCHED", "content size", content.length() + "");
  }

  @Test
  public void test_500KB_ORI() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_500KB);
    String content = docReaderOld.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_500KB_ORI", "ORI", "content size",content.length()+"");
  }

  @Test
  public void test_500KB_PATCHED() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_500KB);
    String content = docReaderNew.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_500KB_PATCHED", "PATCHED", "content size",content.length()+"");
  }

  @Test
  public void test_27MB_formulas_ORI() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_FORMULA);
    String content = docReaderOld.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_27MB_formulas", "ORI", "content size",content.length()+"");
  }

  @Test
  public void test_27MB_formulas_PATCHED() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_FORMULA);
    String content = docReaderNew.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_27MB_formulas", "PATCHED", "content size",content.length()+"");
  }

  @Test
  public void test_17MB_text_ORI() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_TEXT);
    String content = docReaderOld.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_17MB_text", "ORI", "content size",content.length()+"");
  }

  @Test
  public void test_17MB_text_PATCHED() throws IOException, DocumentReadException {
    InputStream docIS = MSExcelDocumentReaderTest.class.getResourceAsStream("/" + MS_XLS_TEXT);
    String content = docReaderNew.getContentAsText(docIS);
    docIS.close();
    addMoreInfos("test_17MB_text", "PATCHED", "content size",content.length()+"");
  }

  private void addMoreInfos(String test, String version, String type, String value) {
    Map<String, Object> versionData;
    if (moreInfos.containsKey(test)) {
      Map<String, Map<String, Object>> testData = moreInfos.get(test);

      if(testData.containsKey(version)) {
        versionData = testData.get(version);
      } else {
        versionData = new TreeMap<String, Object>();
        testData.put(version,versionData);
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

  }

  @AfterClass
  public static void afterTests() {
    for (Map.Entry<String, Map<String, Map<String, Object>>> entryTest : moreInfos.entrySet()) {
      String testName = entryTest.getKey();
      Map<String, Map<String, Object>> testData = entryTest.getValue();
      for (Map.Entry<String, Map<String, Object>> entryVersion : testData.entrySet()) {
        String testVersion = entryVersion.getKey();
        Map<String, Object> testInfos = entryVersion.getValue();
      }
    }
    System.out.println("MORE INFOS:\n"+moreInfos.toString());
  }
}
