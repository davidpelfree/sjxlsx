package com.incesoft.tools.excel.xlsx;

/**
 * @author floyd
 *
 */
public class FontRegion {

	public FontRegion(short start, short end, Font font) {
		super();
		this.start = start;
		this.end = end;
		this.font = font;
	}
	public FontRegion() {
		super();
		// TODO Auto-generated constructor stub
	}
	short start;
	short end;
	Font font;
	short getEnd() {
		return end;
	}
	void setEnd(short end) {
		this.end = end;
	}
	Font getFont() {
		return font;
	}
	void setFont(Font font) {
		this.font = font;
	}
	short getStart() {
		return start;
	}
	void setStart(short start) {
		this.start = start;
	}
}
