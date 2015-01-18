package com.incesoft.tools.excel.xlsx;


import java.util.ArrayList;
import java.util.List;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.XMLStreamWriter;

/**
 * @author floyd
 * 
 * <si> <r><t>afa</t></r> <r><rPr><color rgb="FF0000"/></rPr><t>你好</t></r>
 * <r><t>df</t></r></si>
 */
public class RichText extends SharedStringText {
	public RichText(String text) {
		super(text);
	}

	RichText() {
		super();
	}

	private List<FontRegion> regions;

	public static final int RANGE_ALL_TEXT = -1;

	private void writeTagT(XMLStreamWriter writer, String text, int start,
			int end) throws XMLStreamException {
		writer.writeStartElement("t");
		if (text.charAt(start) == ' ' || text.charAt(end - 1) == ' ')
			writer.writeAttribute("xml:space", "preserve");
		writer.writeCharacters(text.substring(start, end));
		writer.writeEndElement();// end t
	}

	public void serialize(XMLStreamWriter writer) throws XMLStreamException {
		if (text == null || text.length() == 0) {
			throw new IllegalStateException("empty text of rich text,index="
					+ getIndex());
		}
		writer.writeStartElement("si");

		short lastEnd = 0;
		for (FontRegion region : regions) {
			if (region.start == RANGE_ALL_TEXT) {
				// if range all text,there'll be no more region
				region.start = 0;
				region.end = (short) text.length();
			}
			if (lastEnd < region.getStart()) {
				writer.writeStartElement("r");
				writeTagT(writer, text, lastEnd, region.getStart());
				writer.writeEndElement();// end r
			}
			writer.writeStartElement("r");
			region.getFont().serializeAsRichText(writer);
			writeTagT(writer, text, region.getStart(), region.getEnd());
			writer.writeEndElement();// end r

			lastEnd = region.getEnd();
		}
		// TAIL
		if (lastEnd < text.length()) {
			writer.writeStartElement("r");
			writeTagT(writer, text, lastEnd, text.length());
			writer.writeEndElement();// end r
		}
		writer.writeEndElement();
	}

	public void parse(XMLStreamReader reader) {

	}

	public void applyFont(Font font) {
		applyFont(font, RANGE_ALL_TEXT, RANGE_ALL_TEXT);
	}

	public void applyFont(Font font, int start, int end) {
		if (start != RANGE_ALL_TEXT && start >= end) {
			throw new IllegalArgumentException("apply font with start >= end");
		}
		if (regions == null) {
			regions = new ArrayList<FontRegion>(2);
		}
		if (start == RANGE_ALL_TEXT && regions.size() > 0) {
			regions.clear();
		}
		regions.add(new FontRegion((short) start, (short) end, font));
	}

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

	public List<FontRegion> getRegions() {
		return regions;
	}
}
