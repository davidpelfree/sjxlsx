package com.incesoft.tools.excel.xlsx;


import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * @author floyd
 * 
 */
public class Fill extends IndexedObject implements SerializableEntry {
	Fill() {
	}

	String fgColor;

	public String getFgColor() {
		return fgColor;
	}

	public void setFgColor(Integer fgColor) {
		this.fgColor = Integer.toHexString(fgColor);
	}

	public void setFgColor(String fgColor) {
		this.fgColor = fgColor;
	}

	/**
	 * <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00" />
	 * <bgColor indexed="64" /></patternFill></fill>
	 */
	public void serialize(XMLStreamWriter writer) throws XMLStreamException {
		writer.writeStartElement("fill");

		writer.writeStartElement("patternFill");
		writer.writeAttribute("patternType", "solid");

		writer.writeStartElement("fgColor");
		writer.writeAttribute("rgb", fgColor);
		writer.writeEndElement();// end fgColor

		writer.writeEndElement();// end patternFill
		writer.writeEndElement();// end fill
	}

	@Override
	public int hashCode() {
		final int PRIME = 31;
		int result = 1;
		result = PRIME * result + ((fgColor == null) ? 0 : fgColor.hashCode());
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
		final Fill other = (Fill) obj;
		if (fgColor == null) {
			if (other.fgColor != null)
				return false;
		} else if (!fgColor.equals(other.fgColor))
			return false;
		return true;
	}
}
