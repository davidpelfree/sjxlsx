package com.incesoft.tools.excel.xlsx;


import javax.xml.stream.XMLStreamReader;

/**
 * @author floyd
 * 
 */
public interface ParsableEntry {
	public void parse(XMLStreamReader reader);
}
