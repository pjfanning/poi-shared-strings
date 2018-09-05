# poi-shared-strings
Memory efficient Shared Strings Table implementation for POI streaming

https://bz.apache.org/bugzilla/show_bug.cgi?id=61832

There is a [file reading sample](https://github.com/pjfanning/poi-shared-strings-sample).

I will soon add a file writing sample but this [test code](https://github.com/pjfanning/poi-shared-strings/blob/master/src/test/java/com/github/pjfanning/poi/xssf/streaming/TestSXSSFWorkbookWithTempFileSharedStringsTable.java) might help to get you started.

When reading files, use `new TempFileSharedStringsTable(opcPackage, true)` to have the shared strings loaded from the xlsx package.

If you are using the TempFileSharedStringsTable when writing files (eg using [SXSSFWorkbook](https://poi.apache.org/apidocs/org/apache/poi/xssf/streaming/SXSSFWorkbook.html)), then use `new TempFileSharedStringsTable(true)` to create an empty table that you can add shared string entries to.
