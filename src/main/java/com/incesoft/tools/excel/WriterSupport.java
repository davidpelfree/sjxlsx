package com.incesoft.tools.excel;

import java.io.File;
import java.io.OutputStream;

import com.incesoft.tools.excel.support.CellFormat;
import com.incesoft.tools.excel.support.XLSWriterSupport;
import com.incesoft.tools.excel.support.XLSXWriterSupport;

abstract public class WriterSupport {

	public final static int TYPE_XLS = 1;
	public final static int TYPE_XLSX = 2;

	protected File file;

	public void setFile(File file) {
		this.file = file;
	}

	protected OutputStream output;

	public void setOutputStream(OutputStream output) {
		this.output = output;
	}

	abstract protected int getMaxRowNumOfSheet();

	abstract public void open();

	abstract public void createNewSheet();

	abstract public void writeRow(String[] rowData);

	abstract public void writeRow(String[] rowData, CellFormat[] formats);

	abstract public void close();

	public static WriterSupport newInstance(int type, File f) {
		WriterSupport support = null;
		if (type == TYPE_XLSX)
			support = new XLSXWriterSupport();
		else
			support = new XLSWriterSupport();
		support.setFile(f);
		return support;
	}

	public static WriterSupport newInstance(int type, OutputStream outputStream) {
		WriterSupport support = null;
		if (type == TYPE_XLSX)
			support = new XLSXWriterSupport();
		else {
			support = new XLSWriterSupport();
		}
		support.setOutputStream(outputStream);
		return support;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public int getRowPos() {
		return rowpos;
	}

	protected int rowpos = getMaxRowNumOfSheet();

	protected int sheetIndex = -1;

	public void increaseRow() {
		rowpos++;

		if (rowpos > getMaxRowNumOfSheet()) {// 判断是否需要新建一个sheet
			sheetIndex++;
			createNewSheet();
			rowpos = -1;
		}
	}

}
