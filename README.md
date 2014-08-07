test-core-doc-reader
====================

## MSExcelDocumentReader benchmarks

    # Launch various XLS Doc Reader algorithm tests
    mvn clean test -Dtest=*MSExcelDocumentReaderStreamTest
    
    # Launch various XLS Doc Reader algorithm tests and generate JUnitBenchmark report
    mvn clean test -Djub.consumers=CONSOLE,H2 -Djub.db.file=.benchmarks -Dtest=*MSExcelDocumentReaderStreamTest
    