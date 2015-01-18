package com.incesoft.tools.excel.support;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.commons.io.IOUtils;

import com.incesoft.tools.excel.WriterSupport;
import com.incesoft.tools.excel.xlsx.CellStyle;
import com.incesoft.tools.excel.xlsx.Fill;
import com.incesoft.tools.excel.xlsx.Font;
import com.incesoft.tools.excel.xlsx.Sheet;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook;

/**
 * @author floyd
 * 
 */
public class XLSXWriterSupport extends WriterSupport {
	SimpleXLSXWorkbook workbook;

	public void open() {
		if (getClass().getResource("/empty.xlsx") == null) {
			throw new IllegalStateException("no empty.xlsx found in classpath");
		}
		workbook = new SimpleXLSXWorkbook(new File(getClass().getResource("/empty.xlsx").getFile()));
	}

	Sheet sheet;

	protected int getMaxRowNumOfSheet() {
		return Integer.MAX_VALUE / 2;
	}

	public void writeRow(String[] rowData) {
		writeRow(rowData, null);
	}

	public void writeRow(String[] rowData, CellFormat[] formats) {
		for (int col = 0; col < rowData.length; col++) {
			String string = rowData[col];
			if (string == null)
				continue;
			CellFormat format = null;
			if (formats != null && formats.length > 0) {
				for (CellFormat cellFormat : formats) {
					if (cellFormat != null && cellFormat.getCellIndex() == col) {
						format = cellFormat;
						break;
					}
				}
			}
			CellStyle cellStyle = null;
			if (format != null && (format.getBackColor() != -1 || format.getForeColor() != -1)) {
				Font font = null;
				Fill fill = null;
				if (format.getForeColor() != -1) {
					font = workbook.createFont();
					font.setColor(format.getForeColor());
				}
				if (format.getBackColor() != -1) {
					fill = workbook.createFill();
					fill.setFgColor(format.getBackColor());
				}
				cellStyle = workbook.createStyle(font, fill);
			}
			sheet.modify(rowpos, col, string, cellStyle);
		}
	}

	public void close() {
		if (workbook == null)
			return;
		OutputStream fos = null;
		try {
			if (file != null) {
				fos = new FileOutputStream(file);
			} else if (output != null) {
				fos = output;
			} else {
				throw new IllegalStateException("no output specified");
			}
			workbook.commit(fos);
		} catch (Exception e) {
			throw new RuntimeException(e);
		} finally {
			if (fos != null)
				IOUtils.closeQuietly(fos);
			if (workbook != null)
				workbook.close();
		}
	}

	public void createNewSheet() {
		if (sheetIndex > 0) {
			throw new IllegalStateException("only one sheet allowed");
		}
		sheet = workbook.getSheet(sheetIndex, true);
	}
}
