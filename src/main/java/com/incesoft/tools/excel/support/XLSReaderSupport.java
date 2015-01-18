package com.incesoft.tools.excel.support;

import java.io.File;
import java.text.SimpleDateFormat;

import com.incesoft.tools.excel.ExcelRowIterator;
import com.incesoft.tools.excel.ReaderSupport;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;

public class XLSReaderSupport extends ReaderSupport {

	private Workbook workbook;

	private Sheet[] sheets;

	private File inputFile;

	public class XLSObjectIterator implements ExcelRowIterator {
		byte sheetindex;

		int totalSheet;

		int rowPos = -1;

		Sheet theSheet;

		int currentSheetrowcount;

		public void init() {
			sheetindex = 0;
			theSheet = sheets[sheetindex];
			totalSheet = sheets.length;
			currentSheetrowcount = theSheet.getRows();
		}

		public boolean nextRow() {
			rowPos++;
			if (rowPos == currentSheetrowcount) {// 当读取最后一行,如果当前读取的是当前sheet的最后一行
				sheetindex++;
				if (sheetindex < totalSheet) {
					rowPos = 1;// 切换到下一个sheet，的首行
					theSheet = sheets[sheetindex];
					currentSheetrowcount = theSheet.getRows();
					if (rowPos >= currentSheetrowcount) {
						return false;
					}
					return true;
				} else {
					return false;// 所有记录里面的最后一行
				}

			}
			return true;
		}

		public String getCellValue(int col) {
			if (col < 0)
				return null;
			try {
				return trimNull(theSheet.getCell(col, rowPos));
			} catch (RuntimeException e) {
				throw e;
			}
		}

		public byte getSheetIndex() {
			return sheetindex;
		}

		public int getRowPos() {
			return rowPos;
		}

		public int getCellCount() {
			Cell[] row = theSheet.getRow(rowPos);
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
			workbook = Workbook.getWorkbook(inputFile);
			sheets = workbook.getSheets();
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public ExcelRowIterator rowIterator() {
		XLSObjectIterator iterator = new XLSObjectIterator();
		iterator.init();
		return iterator;
	}

	public void close() {
		if (this.workbook != null) {
			this.workbook.close();
			this.workbook = null;
		}
	}

	private final static SimpleDateFormat sdf = new SimpleDateFormat(
			"yyyy-MM-dd");

	private static String trimNull(Cell inputcell) {
		if (inputcell == null)
			return null;

		if (inputcell.getType() == CellType.DATE) {
			DateCell dateCell = (DateCell) inputcell;
			return sdf.format(dateCell.getDate());
		}

		String input = inputcell.getContents().trim();
		if (input.length() == 0)
			return null;
		return input;
	}

	public void setInputFile(File file) {
		this.inputFile = file;
	}

	public Workbook getWorkbook() {
		return workbook;
	}
}
