package com.incesoft.tools.excel.support;

public class CellFormat {
	public CellFormat() {
		super();
	}

	public CellFormat(int foreColor, int backColor, int cellIndex) {
		super();
		this.foreColor = foreColor;
		this.backColor = backColor;
		this.cellIndex = cellIndex;
	}

	private String fontName = "宋体";

	public String getFontName() {
		return fontName;
	}

	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	/**
	 * foreground/word color
	 */
	private int foreColor = -1;
	/**
	 * background color
	 */
	private int backColor = -1;
	/**
	 * the index of cell in the row
	 */
	private int cellIndex;

	public int getCellIndex() {
		return cellIndex;
	}

	public void setCellIndex(int cellIndex) {
		this.cellIndex = cellIndex;
	}

	public int getForeColor() {
		return foreColor;
	}

	public void setForeColor(int foreColor) {
		this.foreColor = foreColor;
	}

	public int getBackColor() {
		return backColor;
	}

	public void setBackColor(int backColor) {
		this.backColor = backColor;
	}

}
