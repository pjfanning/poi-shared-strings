# poi-shared-strings
Memory efficient Shared Strings Table implementation for [POI](https://poi.apache.org/) xlsx streaming. Supports read and write use cases when used with POI 4.0.0.

https://bz.apache.org/bugzilla/show_bug.cgi?id=61832

The `TempFileSharedStringsTable` uses a [H2 MVStore](http://www.h2database.com/html/mvstore.html) to store the Excel Shared String data. The MVStore data can be encrypted using a generated password.

This class can be used instead of the POI [SharedStringsTable](https://poi.apache.org/apidocs/org/apache/poi/xssf/model/SharedStringsTable.html) and [ReadOnlySharedStringsTable](https://poi.apache.org/apidocs/org/apache/poi/xssf/eventusermodel/ReadOnlySharedStringsTable.html). It is only useful if you expect to need to support large numbers of shared string entries.

## Samples

There is a [file reading sample](https://github.com/pjfanning/poi-shared-strings-sample).

I will soon add a file writing sample but this [test code](https://github.com/pjfanning/poi-shared-strings/blob/master/src/test/java/com/github/pjfanning/poi/xssf/streaming/TestSXSSFWorkbookWithTempFileSharedStringsTable.java) might help to get you started.

## Usage

When reading files, use `new TempFileSharedStringsTable(opcPackage, true)` to have the shared strings loaded from the xlsx package.

If you are using the TempFileSharedStringsTable when writing files (eg using [SXSSFWorkbook](https://poi.apache.org/apidocs/org/apache/poi/xssf/streaming/SXSSFWorkbook.html)), then use `new TempFileSharedStringsTable(true)` to create an empty table that you can add shared string entries to.
