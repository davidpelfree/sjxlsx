sjxlsx - the efficient Java API for XLSX
=======================================

It is a **simple and efficient** java tool for reading, writing and modifying XLSX. The most important purpose to code it is for performance consideration -- all the popular ones like POI sucks in both memory consuming and parse/write speed.

- memory

sjxlsx provides two modes (classic & stream) to read/modify sheets. In classic mode, all records of the sheet will be loaded. In stream mode (also named iterate mode), you can read record one after another which save a lot memory.

- speed

Microsoft XLSX use XML+zip (OOXML) to store the data. So, to be fast, sjxlsx use STAX for XML input and output. And I recommend the WSTX implementation of STAX (it's the fastest in my testing).

Sample code
-----------
```
package com.incesoft.cms.util.excel;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

import com.incesoft.cms.util.excel.Sheet.SheetRowReader;
import com.incesoft.cms.util.excel.SimpleXLSXWorkbook.Commiter;

/**
 * @author floyd
 * 
 */
public class TestSJXLSX {

        public static void addStyleAndRichText(SimpleXLSXWorkbook wb, Sheet sheet)
                        throws Exception {
                Font font2 = wb.createFont();
                font2.setColor("FFFF0000");
                Fill fill = wb.createFill();
                fill.setFgColor("FF00FF00");
                CellStyle style = wb.createStyle(font2, fill);

                RichText richText = wb.createRichText();
                richText.setText("test_text");
                Font font = wb.createFont();
                font.setColor("FFFF0000");
                richText.applyFont(font, 1, 2);
                sheet.modify(0, 0, (String) null, style);
                sheet.modify(1, 0, richText, null);
        }

        static public void addRecordsOnTheFly(SimpleXLSXWorkbook wb, Sheet sheet,
                        int rowOffset) {
                int columnCount = 10;
                int rowCount = 10;
                int offset = rowOffset;
                for (int r = offset; r < offset + rowCount; r++) {
                        int modfiedRowLength = sheet.getModfiedRowLength();
                        for (int c = 0; c < columnCount; c++) {
                                sheet.modify(modfiedRowLength, c, r + "," + c, null);
                        }
                }
        }

        private static void printRow(int rowPos, Cell[] row) {
                int cellPos = 0;
                for (Cell cell : row) {
                        System.out.println(Sheet.getCellId(rowPos, cellPos) + "="
                                        + cell.getValue());
                        cellPos++;
                }
        }

        public static void testLoadALL(SimpleXLSXWorkbook workbook) {
                // medium data set,just load all at a time
                Sheet sheetToRead = workbook.getSheet(0);
                List<Cell[]> rows = sheetToRead.getRows();
                int rowPos = 0;
                for (Cell[] row : rows) {
                        printRow(rowPos, row);
                        rowPos++;
                }
        }

        public static void testIterateALL(SimpleXLSXWorkbook workbook) {
                // here we assume that the sheet contains too many rows which will leads
                // to memory overflow;
                // So we get sheet without loading all records
                Sheet sheetToRead = workbook.getSheet(0, false);
                SheetRowReader reader = sheetToRead.newReader();
                Cell[] row;
                int rowPos = 0;
                while ((row = reader.readRow()) != null) {
                        printRow(rowPos, row);
                        rowPos++;
                }
        }

        public static void testWrite(SimpleXLSXWorkbook workbook,
                        OutputStream outputStream) throws Exception {
                Sheet sheet = workbook.getSheet(0);
                addRecordsOnTheFly(workbook, sheet, 0);
                workbook.commit(outputStream);
        }

        /**
         * Commit serveral times for large data set
         * 
         * @param workbook
         * @param output
         * @throws Exception
         */
        public static void testWriteByIncrement(SimpleXLSXWorkbook workbook,
                        OutputStream output) throws Exception {
                Commiter commiter = workbook.newCommiter(output);
                commiter.beginCommit();

                Sheet sheet = workbook.getSheet(0, false);
                commiter.beginCommitSheet(sheet);
                addRecordsOnTheFly(workbook, sheet, 0);
                commiter.commitSheetWrites();
                addRecordsOnTheFly(workbook, sheet, 20);
                commiter.commitSheetWrites();
                addRecordsOnTheFly(workbook, sheet, 40);
                commiter.commitSheetWrites();
                commiter.endCommitSheet();

                commiter.endCommit();
        }

        /**
         * first, modify the original sheet; and then append some data
         * 
         * @param workbook
         * @param output
         * @throws Exception
         */
        public static void testMergeBeforeWrite(SimpleXLSXWorkbook workbook,
                        OutputStream output) throws Exception {
                Sheet sheet = workbook.getSheet(0, false);// assuming original data
                                                                                                        // set is large
                addStyleAndRichText(workbook, sheet);
                addRecordsOnTheFly(workbook, sheet, 5);

                Commiter commiter = workbook.newCommiter(output);
                commiter.beginCommit();
                commiter.beginCommitSheet(sheet);
                // merge it first,otherwise the modification will not take effect
                commiter.commitSheetModifications();

                // row = -1, for appending after the last row
                sheet.modify(-1, 1, "append1", null);
                sheet.modify(-1, 2, "append2", null);
                // lets assume there are many rows here...
                commiter.commitSheetWrites();// flush writes,save memory

                sheet.modify(-1, 1, "append3", null);
                sheet.modify(-1, 2, "append4", null);
                // lets assume there are many rows here,too ...
                commiter.commitSheetWrites();// flush writes,save memory

                commiter.endCommitSheet();
                commiter.endCommit();
        }

        private static SimpleXLSXWorkbook newWorkbook() {
                return new SimpleXLSXWorkbook(new File("/sample.xlsx"));
        }

        private static OutputStream newOutput(String suffix) throws Exception {
                return new BufferedOutputStream(new FileOutputStream("/sample_"
                                + suffix + ".xlsx"));
        }

        public static void main(String[] args) throws Exception {
                SimpleXLSXWorkbook workbook = newWorkbook();
                // READ by classic mdoe - load all records
                testLoadALL(newWorkbook());
                // READ by stream mode - iterate records one by one
                testIterateALL(newWorkbook());

                // WRITE - we take WRITE as a special kind of MODIFY
                OutputStream output = newOutput("write");
                testWrite(workbook, output);
                output.close();

                // WRITE large data
                output = newOutput("write_inc");
                testWriteByIncrement(workbook, output);
                output.close();

                // MODIFY it and WRITE large data
                output = newOutput("merge_write");
                testMergeBeforeWrite(workbook, output);
                output.close();
        }
}
```
