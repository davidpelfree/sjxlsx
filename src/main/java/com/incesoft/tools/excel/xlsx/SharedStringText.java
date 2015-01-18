package com.incesoft.tools.excel.xlsx;


import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * @author floyd
 * 
 */
public class SharedStringText extends IndexedObject implements
		SerializableEntry {

	protected String text;

	public SharedStringText(String text) {
		this.text = text;
	}

	SharedStringText() {
	}

	public void serialize(XMLStreamWriter writer) throws XMLStreamException {
		if (text == null) {
			throw new IllegalStateException("empty text of plain text,index="
					+ getIndex());
		}
		writer.writeStartElement("si");
		writer.writeStartElement("t");
		writer.writeCharacters(text);
		writer.writeEndElement();
		writer.writeEndElement();
	}

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

}
