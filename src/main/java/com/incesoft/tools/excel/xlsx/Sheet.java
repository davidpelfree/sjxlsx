package com.incesoft.tools.excel.xlsx;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.XMLStreamWriter;

import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.ModifyEntry;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.XMLStreamCreator;

/**
 * One Sheet in a workbook.It provides read and write functions of the
 * rows/cells.
 *
 * @author floyd
 *
 */
public class Sheet {

  private int sheetIndex;

  private SimpleXLSXWorkbook workbook;

  public Sheet(int sheetIndex, SimpleXLSXWorkbook workbook) {
    this.sheetIndex = sheetIndex;
    this.workbook = workbook;
  }

  List<Cell[]> parsedRows = new ArrayList<Cell[]>();

  // READ>>>
  /**
   * sheetParser
   */
  XMLStreamReader reader;

  /**
   * convenient methods to load all rows
   *
   */
  boolean alreadyParsed = false;

  void parseAllRows() {
    if (!alreadyParsed) {
      alreadyParsed = true;
      new SheetRowReader(this, workbook.getSheetReader(sheetIndex + 1), true).readRow();
    }
  }

  public static final Cell[] EMPTY_ROW = new Cell[0];

  /**
   * Load row one by one for performance consideration. If using this method
   * with {@link #setAddToMemory(boolean)}=true,you should be careful with the
   * 'modifyXXX' methods,because it won't put any rows into memory -- sth like
   * 'readonly' mode.Besides,the stream-modify api will come in soon.
   *
   * @return null if there is no more rows
   */

  public static class IteratorStatus {
    int rowIndex = -1;

    public int getRowIndex() {
      return rowIndex;
    }
  }

  public static class SheetRowReader {
    public final static int MAX_COLUMN_SPAN = 26 * 26 + 26; // [A-Z]x[A-Z] +
                                                            // [A-Z]

    IteratorStatus status = new IteratorStatus();

    Sheet sheet;

    XMLStreamReader reader;

    boolean loadEagerly;

    SheetRowReader(Sheet sheet, XMLStreamReader reader, boolean loadEagerly) {
      this.sheet = sheet;
      this.reader = reader;
      this.loadEagerly = loadEagerly;
    }

    public IteratorStatus getStatus() {
      return status;
    }

    private int lastRowIndex = -1;

    private Cell[] delayRow;

    private int columnCount = 0;

    /**
     * iterate over the rows(including empty rows)
     *
     * @return null when no more rows.
     */
    public Cell[] readRow() {
      try {
        if (!loadEagerly) {
          status.rowIndex++;// new row
          if (status.rowIndex < lastRowIndex) {// empty rows
            return EMPTY_ROW;
          }
          if (delayRow != null) {
            Cell[] ret = delayRow;
            delayRow = null;
            return ret;
          }
        }

        Cell[] ret = null;
        String r, s = null, t, text;
        String v;
        int columnspan;
        String spans;
        while (reader.hasNext()) {
          int type = reader.next();
          switch (type) {
          case XMLStreamReader.START_ELEMENT:
            if ("col".equals(reader.getLocalName())) {
              increaseColumnCountInSheet();
            }
            if ("row".equals(reader.getLocalName())) {
              spans = reader.getAttributeValue(null, "spans");
              if (spans == null) {
                if (anyColumnsDefinedBeforeSheetData()) {
                  setLastRowIndex();
                  columnCount = checkIfColumnLimitIsReachedIfYesReturnMaxColumn(columnCount);
                  ret = new Cell[columnCount];
                } else {
                  // trully empty row
                  ret = EMPTY_ROW;
                }
              } else {
                setLastRowIndex();
                columnspan = Integer.valueOf(spans.substring(spans.indexOf(":") + 1));
                columnspan = checkIfColumnLimitIsReachedIfYesReturnMaxColumn(columnspan);
                ret = new Cell[columnspan];
              }
            } else if ("c".equals(reader.getLocalName())) {
              if (ret != null) {
                t = reader.getAttributeValue(null, "t");
                // s = reader.getAttributeValue(null, "s");
                r = reader.getAttributeValue(null, "r");
                text = null;
                v = null;
                while (reader.hasNext()) {
                  type = reader.next();
                  if (type == XMLStreamReader.CHARACTERS) {
                    v = reader.getText();
                    if ("s".equals(t)) {
                      text = sheet.workbook.getSharedStringValue(Integer.valueOf(v));
                    }
                  } else if (type == XMLStreamReader.END_ELEMENT && "c".equals(reader.getLocalName())) {
                    break;
                  }
                }
                if (r.charAt(1) < 'A') {// number
                  ret[r.charAt(0) - 'A'] = new Cell(r, s, t, v, text);
                } else if (r.length() > 2 && r.charAt(2) < 'A') {
                  int i = (r.charAt(1) - 'A') + (r.charAt(0) - 'A' + 1) * 26;
                  if (i < MAX_COLUMN_SPAN)
                    ret[i] = new Cell(r, s, t, v, text);
                }
                // ignore columns larger than AZ
              } else {
                throw new IllegalStateException("<c> mal-format");
              }
            }
            break;
          case XMLStreamReader.END_ELEMENT:
            if ("row".equals(reader.getLocalName())) {
              if (loadEagerly) {
                status.rowIndex++;
                if (status.rowIndex < lastRowIndex) {
                  if (sheet.addToMemory) {
                    // fill the empty rows
                    for (int i = 0; i < lastRowIndex - status.rowIndex; i++) {
                      sheet.parsedRows.add(EMPTY_ROW);
                    }
                  }
                }
                status.rowIndex = lastRowIndex;
              }
              if (sheet.addToMemory) {
                sheet.parsedRows.add(ret);
              }
              if (loadEagerly) {
                ret = null;
              } else if (status.rowIndex < lastRowIndex) {
                delayRow = ret;
                return EMPTY_ROW;
              } else {
                return ret;
              }
            }
            break;
          default:
            break;
          }
        }
      } catch (XMLStreamException e) {
        throw new RuntimeException(e);
      }
      return null;
    }

    private boolean anyColumnsDefinedBeforeSheetData() {
      return columnCount > 0;
    }

    private int checkIfColumnLimitIsReachedIfYesReturnMaxColumn(int columnspan) {
      if (columnspan > MAX_COLUMN_SPAN) {
        columnspan = MAX_COLUMN_SPAN;
      }
      return columnspan;
    }

    private void setLastRowIndex() {
      lastRowIndex = Integer.valueOf(reader.getAttributeValue(null, "r")) - 1;
    }

    private void increaseColumnCountInSheet() {
      columnCount++;
    }

    public void remove() {
      throw new UnsupportedOperationException();
    }

  }

  public SheetRowReader newReader() {
    return new SheetRowReader(this, workbook.getSheetReader(sheetIndex + 1), false);
  }

  private boolean addToMemory = true;

  /**
   * Should only be called in one loop while termination condition is nextRow()
   * == null
   *
   * @return
   */
  public List<Cell[]> getRows() {
    if (!(alreadyParsed && addToMemory)) {
      throw new IllegalStateException("rows not parsed,it should only be used in classic mode");
    }
    return parsedRows;
  }

  private int rowCount = -2;

  // count of the lazy or non-lazy rows
  public int getRowCount() {
    if (alreadyParsed && addToMemory) {
      return parsedRows.size();
    }
    if (rowCount == -2) {
      rowCount = 0;
      XMLStreamReader reader = workbook.getSheetReader(sheetIndex + 1);
      try {
        // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        // <worksheet ...><dimension ref="A1:C3"/>...<sheetData>
        loopR: while (reader.hasNext()) {
          int type = reader.next();
          switch (type) {
          case XMLStreamReader.START_ELEMENT:
            if ("dimension".equals(reader.getLocalName())) {
              String v = reader.getAttributeValue(null, "ref");
              if (v != null) {
                String[] spanPair = v.replaceAll("[A-Z]", "").split(":");
                if (spanPair.length == 2) {
                  try {
                    rowCount = Integer.valueOf(spanPair[1]) - Integer.valueOf(spanPair[0]) + 1;
                  } catch (NumberFormatException e) {
                  }
                  break loopR;
                }
              }
            } else if ("row".equals(reader.getLocalName())) {
              int r = Integer.valueOf(reader.getAttributeValue(null, "r"));
              if (r > rowCount)
                rowCount = r;
            }
            break;
          }
        }
      } catch (XMLStreamException e) {
        throw new RuntimeException(e);
      } finally {
        try {
          reader.close();
        } catch (XMLStreamException e) {
        }
      }
    }
    return rowCount;
  }

  public void clearRows() {
    parsedRows.clear();
  }

  public String getCellValue(int row, int column) {
    if (!(alreadyParsed && addToMemory)) {
      throw new IllegalStateException("rows not parsed,it should only be used in classic mode");
    }
    if (row < parsedRows.size()) {
      Cell[] rowEntry = parsedRows.get(row);
      if (rowEntry == EMPTY_ROW) {
        return null;
      }
      if (column < rowEntry.length) {
        return rowEntry[column] == null ? null : rowEntry[column].getValue();
      }
      return null;
    }
    return null;
  }

  public String getCellValue(String cellId) {
    if (!(alreadyParsed && addToMemory)) {
      throw new IllegalStateException("rows not parsed,it should only be used in classic mode");
    }
    return getCellValue(Integer.valueOf(cellId.substring(1)) - 1, cellId.charAt(0) - 'A');
  }

  // READ<<<

  // MODIFY>>>
  int modifiedRowLength = 0;

  int lastCommittedRowLength = 0;

  public boolean isModified() {
    // more than once commit/merge
    return lastCommittedRowLength > 0 || modifiedRowLength > 0;
  }

  /**
   * {rowIndex,modifications}
   */
  HashMap<Integer, List<ModifyEntry>> modifications = new HashMap<Integer, List<ModifyEntry>>(100);

  /**
   * max row index have been modified - for append/add convenience
   *
   * @return
   */
  public int getModfiedRowLength() {
    return modifiedRowLength == 0 ? lastCommittedRowLength : modifiedRowLength;
  }

  public int modify(int r, int c, String comment) {
    if (c > SheetRowReader.MAX_COLUMN_SPAN - 1) {
      throw new IllegalArgumentException("column index(" + c + ") exceeded the limit(" + (SheetRowReader.MAX_COLUMN_SPAN - 1) + ")");
    }
    if (r == -1) {
      r = getModfiedRowLength();
    } else if (lastCommittedRowLength > 0 && r < lastCommittedRowLength) {
      throw new IllegalStateException("after merge,only add allowed");
    }
    // resize
    if (r >= modifiedRowLength) {
      modifiedRowLength = r + 1;
    }
    // modify
    List<ModifyEntry> modified = modifications.get(r);
    if (modified == null) {
      modified = new ArrayList<ModifyEntry>();
      modifications.put(r, modified);
    }
    modified.add(new ModifyEntry(r, c, null, null, comment));
    return r;
  }

  /**
   *
   * @param r
   *          -1 to append
   * @param c
   * @param text
   * @param style
   */
  public int modify(int r, int c, String text, CellStyle style) {
    if (c > SheetRowReader.MAX_COLUMN_SPAN - 1) {
      throw new IllegalArgumentException("column index(" + c + ") exceeded the limit(" + (SheetRowReader.MAX_COLUMN_SPAN - 1) + ")");
    }
    if (r == -1) {
      r = getModfiedRowLength();
    } else if (lastCommittedRowLength > 0 && r < lastCommittedRowLength) {
      throw new IllegalStateException("after merge,only add allowed");
    }
    // resize
    if (r >= modifiedRowLength) {
      modifiedRowLength = r + 1;
    }
    // modify
    List<ModifyEntry> modified = modifications.get(r);
    if (modified == null) {
      modified = new ArrayList<ModifyEntry>();
      modifications.put(r, modified);
    }
    SharedStringText t = null;
    if (text != null) {
      t = workbook.createPlainText(text);
      t.setText(text);
    }
    modified.add(new ModifyEntry(r, c, t, style));
    return r;
  }

  public int modify(int r, int c, RichText text, CellStyle style) {
    if (c > SheetRowReader.MAX_COLUMN_SPAN - 1) {
      throw new IllegalArgumentException("column index(" + c + ") exceeded the limit(" + (SheetRowReader.MAX_COLUMN_SPAN - 1) + ")");
    }
    if (r == -1) {
      r = getModfiedRowLength();
    } else if (lastCommittedRowLength > 0 && r < lastCommittedRowLength) {
      throw new IllegalStateException("after merge,only add allowed");
    }
    // resize
    if (r >= modifiedRowLength) {
      modifiedRowLength = r + 1;
    }
    // modify
    List<ModifyEntry> modified = modifications.get(r);
    if (modified == null) {
      modified = new ArrayList<ModifyEntry>();
      modifications.put(r, modified);
    }
    modified.add(new ModifyEntry(r, c, text, style));
    return r;
  }

  public void writeDocumentStart(XMLStreamWriter writer) throws XMLStreamException {
    // PATCH ,cause XmlStreamRead's START_DOCUMENT event never occurred
    writer.writeStartDocument("UTF-8", "1.0");
  }

  private Cell modifyCellInternal(ModifyEntry modification, Cell cell) {
    if (modification != null) {
      if (cell == null) {
        cell = new Cell();
      }
      // if text changed or region font applied
      SharedStringText text = modification.text;
      if (text != null) {
        if (text instanceof RichText) {
          if (text.getText() == null && cell != null) {
            if (cell.getValue() == null) {
              throw new IllegalStateException("there is no cell content for richtext modification,cell="
                  + getCellId(modification.r, modification.c));
            }
            text.setText(cell.getValue());
          }
        }
        cell.setV(String.valueOf(text.getIndex()));
        cell.setT("s");
      }
      if (modification.style != null)
        cell.setS(String.valueOf(modification.style.index));
      if (modification.comment != null)
        cell.setComment(modification.comment);
    }
    return cell;
  }

  private class SheetWriter {
    XMLStreamWriter xmlWriter;

    int rowIndex = -1;

    SheetCommentWriter commentWriter;

    public SheetWriter(XMLStreamWriter xmlWriter, SheetCommentWriter commentWriter) {
      super();
      this.xmlWriter = xmlWriter;
      this.commentWriter = commentWriter;
    }

    void writeStart() throws XMLStreamException {
      xmlWriter.writeStartDocument("UTF-8", "1.0");
      xmlWriter.writeStartElement("worksheet");
      xmlWriter.writeNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
      xmlWriter.writeNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
      xmlWriter.writeStartElement("sheetData");
    }

    void writeEnd(String commentRId, String vmlRid, String settingRid) throws XMLStreamException {
      xmlWriter.writeEndElement();// end sheetData
      if (settingRid != null) {
        xmlWriter.writeStartElement("pageSetup");
        xmlWriter.writeAttribute("r:id", settingRid);
        xmlWriter.writeEndElement();
      }
      if (commentWriter.startWriten) {
        // refer to the vml drawing -- the design sucks
        // <legacyDrawing r:id="rId3" />
        if (vmlRid == null)
          throw new IllegalStateException("vmlRid should not be null");
        xmlWriter.writeStartElement("legacyDrawing");
        xmlWriter.writeAttribute("r:id", vmlRid);
        xmlWriter.writeEndElement();
      }
      xmlWriter.writeEndElement();// end worksheet
      // end the comment if necessary
      commentWriter.writeEnd();
    }

    private Cell cell;

    void writeRow(Cell[] row, int rowIndex) throws XMLStreamException {
      this.rowIndex = rowIndex;
      int maxCol = 0;// max column index
      List<ModifyEntry> modificationList = modifications.isEmpty() ? null : modifications.get(rowIndex);
      if (modificationList != null) {
        for (ModifyEntry entry : modificationList) {
          if (entry.c >= maxCol) {
            maxCol = entry.c;
          }
        }
      } else if (row == EMPTY_ROW) {
        // no modfication and new raw rowdata
        return;
      }

      // resize the row if necessary
      if (row == EMPTY_ROW) {
        row = new Cell[maxCol + 1];
      } else if (maxCol > row.length - 1) {
        Cell[] newrow = new Cell[maxCol + 1];
        System.arraycopy(row, 0, newrow, 0, row.length);
        row = newrow;
      }

      /**
       * <row r="1" spans="1:2">
       */
      xmlWriter.writeStartElement("row");
      xmlWriter.writeAttribute("r", String.valueOf(rowIndex + 1));
      xmlWriter.writeAttribute("spans", "1:" + row.length);

      // modifiy cells
      if (modificationList != null) {
        for (ModifyEntry modification : modificationList) {
          row[modification.c] = modifyCellInternal(modification, row[modification.c]);
        }
      }

      // serialize the row
      for (int col = 0; col < row.length; col++) {
        cell = row[col];
        if (cell == null)
          continue;

        if (cell.getR() == null) {
          cell.setR(getCellId(rowIndex, col));
        }
        writeCell(rowIndex, col, cell, xmlWriter);
      }

      xmlWriter.writeEndElement();// end row
    }

    private void writeCell(int row, int c, Cell cell, XMLStreamWriter writer) throws XMLStreamException {
      if (cell != null) {
        writer.writeStartElement("c");
        writer.writeAttribute("r", cell.getR());
        if (cell.getS() != null) {
          writer.writeAttribute("s", cell.getS());
        }
        if (cell.getT() != null) {
          writer.writeAttribute("t", cell.getT());
        }
        if (cell.getV() != null) {
          writer.writeStartElement("v");
          writer.writeCharacters(String.valueOf(cell.getV()));
          writer.writeEndElement();
        }
        if (cell.getComment() != null)
          commentWriter.writeComment(row, c, cell.getComment());
        writer.writeEndElement();// end c
      }
    }
  }

  private SheetWriter sheetsWriter;

  void writeSheetStart(XMLStreamWriter writer, XMLStreamCreator commentWriter, XMLStreamCreator vmlWriter) throws XMLStreamException {
    if (sheetsWriter != null) {
      throw new IllegalStateException("sheets can only be merged once.");
    }
    sheetsWriter = new SheetWriter(writer, new SheetCommentWriter(commentWriter, vmlWriter));
    sheetsWriter.writeStart();
  }

  public boolean isCommentModified() {
    return this.sheetsWriter != null && this.sheetsWriter.commentWriter != null && this.sheetsWriter.commentWriter.startWriten;
  }

  void writeSheetEnd(String commentRId, String vmlRid, String settingRid) throws XMLStreamException {
    sheetsWriter.writeEnd(commentRId, vmlRid, settingRid);
  }

  void writeSheet() throws XMLStreamException {
    merged = true;
    // after merged & clean up,just append
    for (int rowIndex = sheetsWriter.rowIndex; rowIndex < modifiedRowLength; rowIndex++) {
      sheetsWriter.writeRow(EMPTY_ROW, rowIndex);
    }
    cleanUp();
  }

  private boolean merged = false;

  /**
   * <?xml version="1.0" encoding="UTF-8" standalone="yes"?> <worksheet
   * xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r=
   * "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
   * ><sheetData>
   *
   * @param writer
   * @throws XMLStreamException
   */
  void mergeSheet() throws XMLStreamException {
    if (merged) {
      writeSheet();
      return;
    }

    merged = true;
    if (alreadyParsed) {
      int rowLen = Math.max(parsedRows.size(), modifiedRowLength);
      for (int rowIndex = 0; rowIndex < rowLen; rowIndex++) {
        if (rowIndex < parsedRows.size()) {
          sheetsWriter.writeRow(parsedRows.get(rowIndex), rowIndex);
        } else {
          sheetsWriter.writeRow(EMPTY_ROW, rowIndex);
        }
      }
    } else {
      SheetRowReader rowReader = newReader();
      Cell[] row = null;
      int rowIndex = -1;
      while ((row = rowReader.readRow()) != null) {
        // check replace rows
        rowIndex = rowReader.getStatus().getRowIndex();
        sheetsWriter.writeRow(row, rowIndex);
      }
      // tail - add
      if (modifiedRowLength - 1 > rowIndex) {
        for (int i = rowIndex + 1; i < modifiedRowLength; i++) {
          sheetsWriter.writeRow(EMPTY_ROW, i);
        }
      }
    }
    cleanUp();
  }

  // MODFIY<<<

  public void cleanUp() {
    if (sheetsWriter != null)
      lastCommittedRowLength = sheetsWriter.rowIndex + 1;
    modifiedRowLength = 0;
    parsedRows.clear();
    modifications.clear();
  }

  // UTILS>>>

  static private char[] COLUMNS;
  static {
    COLUMNS = new char[26];
    for (int i = 0; i < COLUMNS.length; i++) {
      COLUMNS[i] = (char) (i + (int) 'A');
    }
  }

  public static String getCellId(int r, int c) {
    if (c <= 25) {
      return COLUMNS[c] + String.valueOf(r + 1);
    } else {
      return String.valueOf(new char[] { COLUMNS[c / COLUMNS.length - 1], COLUMNS[c % COLUMNS.length] }) + String.valueOf(r + 1);
    }
  }

  // UTILS<<<

  void setWorkbook(SimpleXLSXWorkbook workbook) {
    this.workbook = workbook;
  }

  public void setAddToMemory(boolean addToMemory) {
    this.addToMemory = addToMemory;
  }

  public int getSheetIndex() {
    return sheetIndex;
  }

  public boolean isMerged() {
    return merged;
  }

}