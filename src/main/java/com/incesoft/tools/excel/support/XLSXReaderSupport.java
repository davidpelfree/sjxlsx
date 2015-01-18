package com.incesoft.tools.excel.support;

import java.io.File;

import com.incesoft.tools.excel.ExcelRowIterator;
import com.incesoft.tools.excel.ReaderSupport;
import com.incesoft.tools.excel.xlsx.Cell;
import com.incesoft.tools.excel.xlsx.Sheet;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;
import com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader;

public class XLSXReaderSupport extends ReaderSupport {

	private SimpleXLSXWorkbook wb;

	private Sheet sheet;

	public XLSXReaderSupport() {
	}

	private File inputFile;

	private boolean lazy = true;

	protected class LazyXLSXObjectIterator implements ExcelRowIterator {

		public int getCellCount() {
			return curRow != null ? curRow.length : 0;
		}

		public String getCellValue(int col) {
			if (curRow == null || col > curRow.length - 1
					|| curRow[col] == null)
				return null;
			String v = curRow[col].getValue();
			return v == null || v.trim().length() == 0 ? null : v.trim();
		}

		public int getRowPos() {
			return reader.getStatus().getRowIndex() - (nextRow != null ? 1 : 0);
		}

		public byte getSheetIndex() {
			return (byte) sheet.getSheetIndex();
		}

		SheetRowReader reader;

		public void init() {
			reader = sheet.newReader();
		}

		private Cell[] curRow;

		private Cell[] lastRow;

		private Cell[] nextRow;

		public boolean nextRow() {
			if (nextRow != null) {
				lastRow = curRow;
				curRow = nextRow;
				nextRow = null;
			} else {
				if (curRow != null) {
					lastRow = curRow;
				}
				curRow = reader.readRow();
			}
			return curRow != null;
		}

		public void prevRow() {
			if (nextRow == null) {
				nextRow = curRow;
				curRow = lastRow;
				lastRow = null;
			}
		}
	}

	protected class XLSXObjectIterator implements ExcelRowIterator {

		int currentSheetRowCount;

		int rowPos = -1;

		public void init() {
			currentSheetRowCount = sheet.getRowCount();
		}

		public boolean nextRow() {
			rowPos++;
			if (rowPos == currentSheetRowCount) {// 当读取最后一行,如果当前读取的是当前sheet的最后一行
				return false;// 所有记录里面的最后一行
			}
			return true;
		}

		public String getCellValue(int col) {
			if (col < 0)
				return null;
			String v = sheet.getCellValue(rowPos, col);
			return v == null || v.trim().length() == 0 ? null : v.trim();
		}

		public byte getSheetIndex() {
			return (byte) sheet.getSheetIndex();
		}

		public int getRowPos() {
			return rowPos;
		}

		public int getCellCount() {
			Cell[] row = sheet.getRows().get(rowPos);
			return row == null ? 0 : row.length;
		}

		public void prevRow() {
			rowPos--;
			if (rowPos == -1) {
				rowPos = 0;
			}
		}
	}

	public void open() {
		try {
			if (!inputFile.exists()) {
				throw new IllegalStateException("not found file "
						+ inputFile.getAbsoluteFile());
			}
			wb = new SimpleXLSXWorkbook(inputFile);
			sheet = wb.getSheet(0, !lazy);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public ExcelRowIterator rowIterator() {
		ExcelRowIterator iterator = lazy ? new LazyXLSXObjectIterator()
				: new XLSXObjectIterator();
		iterator.init();
		return iterator;
	}

	public void close() {
		if (sheet != null) {
			sheet.cleanUp();
			sheet = null;
		}
		if (this.wb != null) {
			this.wb.close();
			this.wb = null;
		}
	}

	public void setInputFile(File file) {
		this.inputFile = file;
	}

	public SimpleXLSXWorkbook getWorkbook() {
		return wb;
	}

	public void setLazy(boolean lazy) {
		this.lazy = lazy;
	}

	public boolean isLazy() {
		return lazy;
	}

}