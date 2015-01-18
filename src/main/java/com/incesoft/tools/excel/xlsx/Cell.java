package com.incesoft.tools.excel.xlsx;


/**
 * @author floyd
 * 
 */
public class Cell {
	Cell(String r, String s, String t, String v, String text) {
		this.text = text;
		this.v = v;
		this.r = r;
		this.s = s;
		this.t = t;
	}

	Cell() {
		super();
	}

	private String text;

	private String v;

	private String r;

	private String s;

	private String t;

	private String comment;

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	String getR() {
		return r;
	}

	void setR(String r) {
		this.r = r;
	}

	String getS() {
		return s;
	}

	void setS(String s) {
		this.s = s;
	}

	String getT() {
		return t;
	}

	void setT(String t) {
		this.t = t;
	}

	String getText() {
		return text;
	}

	void setText(String text) {
		this.text = text;
	}

	String getV() {
		return v;
	}

	void setV(String v) {
		this.v = v;
	}

	public String getValue() {
		if (text != null)
			return text;
		if (v != null)
			return v;
		else
			return null;
	}

	public String toString() {
		return getValue();
	}
}
