package com.incesoft.tools.excel.xlsx;


import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * Font registered in styles.xml
 * 
 * @author floyd
 * 
 */
public class CellStyle extends IndexedObject implements SerializableEntry {
	Font font;

	Fill fill;

	public Fill getFill() {
		return fill;
	}

	void setFill(Fill fill) {
		this.fill = fill;
	}

	public Font getFont() {
		return font;
	}

	void setFont(Font font) {
		this.font = font;
	}

	/**
	 * <cellXfs count="1"> <xf numFmtId="0" fontId="0" fillId="0" borderId="0"
	 * xfId="0"> <alignment vertical="center" /> </xf> </cellXfs>
	 */
	public void serialize(XMLStreamWriter writer) throws XMLStreamException {
		writer.writeStartElement("xf");
		writer.writeAttribute("numFmtId", "0");
		if (font != null)
			writer.writeAttribute("fontId", String.valueOf(font.getIndex()));
		if (fill != null)
			writer.writeAttribute("fillId", String.valueOf(fill.getIndex()));
		writer.writeAttribute("borderId", "0");
		writer.writeAttribute("xfId", "0");

		writer.writeStartElement("alignment");
		writer.writeAttribute("vertical", "center");
		writer.writeEndElement();// end alignment

		writer.writeEndElement();// end xf
	}

	@Override
	public int hashCode() {
		final int PRIME = 31;
		int result = 1;
		result = PRIME * result + ((fill == null) ? 0 : fill.hashCode());
		result = PRIME * result + ((font == null) ? 0 : font.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		final CellStyle other = (CellStyle) obj;
		if (fill == null) {
			if (other.fill != null)
				return false;
		} else if (!fill.equals(other.fill))
			return false;
		if (font == null) {
			if (other.font != null)
				return false;
		} else if (!font.equals(other.font))
			return false;
		return true;
	}

}
