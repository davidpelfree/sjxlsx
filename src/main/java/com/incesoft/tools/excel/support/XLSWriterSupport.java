package com.incesoft.tools.excel.support;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.RGB;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.incesoft.tools.excel.WriterSupport;

/**
 * @author floyd
 * 
 */
public class XLSWriterSupport extends WriterSupport {
	WritableSheet sheet;

	WritableWorkbook workbook;

	public void open() {
		try {
			if (file != null) {
				workbook = Workbook.createWorkbook(file);
			} else if (output != null) {
				workbook = Workbook.createWorkbook(output);
			} else {
				throw new IllegalStateException("no output specified");
			}
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	private static int getRGB(Colour c) {
		RGB defaultRGB = c.getDefaultRGB();
		return defaultRGB.getRed() << 16 | defaultRGB.getGreen() << 8 | defaultRGB.getBlue();
	}

	public static Colour transformColor(int c) {
		int result = Collections.binarySearch(colours, c & 0x00ffffff);
		if (result < 0) {
			result = -(result + 1);
			if (result > colours.size() - 1) {
				result = colours.size() - 1;
			}
		}
		return colourMap.get(colours.get(result));
	}

	static List<Integer> colours = new ArrayList<Integer>();
	static Map<Integer, Colour> colourMap = new HashMap<Integer, Colour>();
	static {
		for (Colour c : Colour.getAllColours()) {
			colourMap.put(getRGB(c), c);
			colours.add(getRGB(c));
		}
		Collections.sort(colours);
	}

	private static final Log log = LogFactory.getLog(XLSWriterSupport.class);

	@Override
	public void writeRow(String[] rowData, CellFormat[] formats) {
		for (int col = 0; col < rowData.length; col++) {
			String cellString = rowData[col];
			if (cellString != null) {
				CellFormat format = null;
				if (formats != null && formats.length > 0) {
					for (CellFormat cellFormat : formats) {
						if (cellFormat != null && cellFormat.getCellIndex() == col) {
							format = cellFormat;
							break;
						}
					}
				}
				Label label = new Label(col, rowpos, cellString);
				try {
					sheet.addCell(label);
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
				if (format != null) {
					WritableCell c = workbook.getSheet(0).getWritableCell(col, rowpos);
					WritableCellFormat newFormat = c.getCellFormat() == null ? new WritableCellFormat()
							: new WritableCellFormat(c.getCellFormat());
					if (format != null && (format.getBackColor() != -1 || format.getForeColor() != -1)) {
						if (format.getBackColor() != -1) {
							try {
								newFormat.setBackground(transformColor(format.getBackColor()));
							} catch (WriteException e) {
								log.error("", e);
							}
						}
						if (format.getForeColor() != -1) {
							try {
								WritableFont writableFont = new WritableFont(WritableFont.createFont(format
										.getFontName()));
								writableFont.setColour(Colour.PINK2);
								newFormat.setFont(writableFont);
							} catch (WriteException e) {
								log.error("", e);
							}
						}
					}
					c.setCellFormat(newFormat);
				}
			}
		}
	}

	public void writeRow(String[] rowData) {
		writeRow(rowData, null);
	}

	public void close() {
		try {
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public void createNewSheet() {
		sheet = workbook.createSheet("sheet" + sheetIndex + 1, sheetIndex);
	}

	public int getMaxRowNumOfSheet() {
		return 60000;
	}
}
