package com.incesoft.tools.excel.xlsx;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * partial font of RichText OR font registered in sharedStrings(index > 0)
 * 
 * @author floyd
 * 
 */
public class Font extends IndexedObject implements SerializableEntry {
	Font() {
	}

	String color;

	// short sz;
	// String fontName;
	// String family;
	// String charset;

	public void setColor(int color) {
		this.color = Integer.toHexString(color);
	}

	public String getColor() {
		return color;
	}

	public void setColor(String color) {
		this.color = color;
	}

	/**
	 * <rPr><color rgb="FF0000"/></rPr>
	 * 
	 * <font><sz val="11"/><color rgb="FFFF0000"/><name val="宋体"/><family
	 * val="2"/><charset val="134"/><scheme val="minor"/></font>
	 */
	public void serialize(XMLStreamWriter writer) throws XMLStreamException {
		writer.writeStartElement("font");
		writer.writeStartElement("color");
		writer.writeAttribute("rgb", color);
		writer.writeEndElement();
		writer.writeEndElement();
	}

	public void serializeAsRichText(XMLStreamWriter writer)
			throws XMLStreamException {
		writer.writeStartElement("rPr");
		writer.writeStartElement("color");
		writer.writeAttribute("rgb", color);
		writer.writeEndElement();
		writer.writeEndElement();
	}

	@Override
	public int hashCode() {
		final int PRIME = 31;
		int result = 1;
		result = PRIME * result + ((color == null) ? 0 : color.hashCode());
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
		final Font other = (Font) obj;
		if (color == null) {
			if (other.color != null)
				return false;
		} else if (!color.equals(other.color))
			return false;
		return true;
	}

}
