package com.incesoft.tools.excel;

import java.io.File;

import com.incesoft.tools.excel.support.XLSReaderSupport;
import com.incesoft.tools.excel.support.XLSXReaderSupport;

abstract public class ReaderSupport {

	public final static int TYPE_XLS = 1;
	public final static int TYPE_XLSX = 2;

	abstract public void setInputFile(File file);

	abstract public void open();

	abstract public ExcelRowIterator rowIterator();

	abstract public void close();

	public static ReaderSupport newInstance(int type, File f) {
		ReaderSupport support = null;
		if (type == TYPE_XLSX)
			support = new XLSXReaderSupport();
		else
			support = new XLSReaderSupport();
		support.setInputFile(f);
		return support;
	}

}
