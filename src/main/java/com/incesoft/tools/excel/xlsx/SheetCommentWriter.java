package com.incesoft.tools.excel.xlsx;


import java.util.StringTokenizer;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.commons.lang.StringUtils;

import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.XMLStreamCreator;


/**
 * 
 * @author floyd
 * 
 */
public class SheetCommentWriter {

	XMLStreamWriter commentsWriter;

	XMLStreamWriter vmlWriter;

	XMLStreamCreator commentWriterCreator;

	XMLStreamCreator vmlWriterCreator;

	public SheetCommentWriter(XMLStreamCreator commentWriterCreator,
			XMLStreamCreator vmlWriterCreator) {
		super();
		this.commentWriterCreator = commentWriterCreator;
		this.vmlWriterCreator = vmlWriterCreator;
	}

	public SheetCommentWriter() {
	}

	public void writeStart() throws XMLStreamException {
		if (commentsWriter == null) {
			commentsWriter = commentWriterCreator.createWriter();
			vmlWriter = vmlWriterCreator.createWriter();
		}
		// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		// <comments
		// xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		// <authors><author>作者</author></authors>
		// <commentList></commentList></comments>
		commentsWriter.writeStartDocument("UTF-8", "1.0");
		commentsWriter.writeStartElement("comments");
		commentsWriter.writeNamespace("xmlns",
				"http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		commentsWriter.writeStartElement("authors");
		commentsWriter.writeStartElement("author");
		commentsWriter.writeCharacters("sjxlsx");
		commentsWriter.writeEndElement();// end author
		commentsWriter.writeEndElement();// end authors
		commentsWriter.writeStartElement("commentList");

		// <xml xmlns:v="urn:schemas-microsoft-com:vml"
		// xmlns:o="urn:schemas-microsoft-com:office:office"
		// xmlns:x="urn:schemas-microsoft-com:office:excel">
		vmlWriter.writeStartElement("xml");
		vmlWriter.writeNamespace("xmlns",
				"http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		vmlWriter.writeNamespace("v", "urn:schemas-microsoft-com:vml");
		vmlWriter
				.writeNamespace("o", "urn:schemas-microsoft-com:office:office");
		vmlWriter.writeNamespace("x", "urn:schemas-microsoft-com:office:excel");
		// <v:shapetype id="_x0000_m1027" coordsize="21600,21600"
		// path="m,l,21600r21600,l21600,xe">
		// </v:shapetype>
		// [VML REFERENCE]->http://www.w3.org/TR/NOTE-VML#_Toc416858391
		vmlWriter.writeStartElement("v:shapetype");
		vmlWriter.writeAttribute("id", this.shapeTypeId);
		vmlWriter.writeAttribute("coordsize", "21600,21600");
		vmlWriter.writeAttribute("path", "m,l,21600r21600,l21600,xe");
		vmlWriter.writeEndElement();// end v:shapetype
	}

	public void writeEnd() throws XMLStreamException {
		if (startWriten) {
			commentsWriter.writeEndElement();// end commentList
			commentsWriter.writeEndElement();// end comments
			vmlWriter.writeEndElement();// end xml
		}
	}

	private String shapeTypeId = "_x0000_m1027";

	private String shapeTypeRefId = "#" + this.shapeTypeId;

	boolean startWriten = false;

	// <comment ref="A1" authorId="0"><text><r><t xml:space="preserve">aaasdf
	// sdf adsf df </t></r></text></comment>

	// <v:shape type="#_x0000_m1027" style='position:absolute;
	// margin-left:108.75pt;margin-top:49.5pt;width:204.75pt;height:90.75pt;
	// z-index:2;mso-wrap-style:tight' fillcolor="#ffffe1" o:insetmode="auto">
	// <x:ClientData ObjectType="Note">
	// <x:Row>0</x:Row>
	// <x:Column>0</x:Column>
	// <x:Anchor>1, 15, 0, 15, 3, 18, 3, 15</x:Anchor>
	// </x:ClientData>
	// </v:shape>
	/**
	 * <pre>
	 * My Explaination:
	 * x:Anchor=1, 15, 0, 16, 3, 3, 4, 18
	 * --- the left&amp;top point accordinate---
	 * 1 - the cell(x) relative to,15 - the x offset relative to cell(Y) in pixel;
	 * 0 - the cell(y) relative to,16 - the y offset relative to cell(X) in pixel;
	 * --- the right&amp;bottom point accordinate---
	 * 3 - the cell(x) relative to,3 - the x offset relative to cell(Y) in pixel;
	 * 4 - the cell(y) relative to,18 - the y offset relative to cell(X) in pixel;
	 * </pre>
	 * 
	 * @param r
	 * @param c
	 * @param comment
	 * @throws XMLStreamException
	 */
	public void writeComment(int r, int c, String comment)
			throws XMLStreamException {
		if (!startWriten) {
			startWriten = true;
			writeStart();
		}
		// write comments
		commentsWriter.writeStartElement("comment");
		commentsWriter.writeAttribute("ref", Sheet.getCellId(r, c));
		commentsWriter.writeAttribute("authorId", "0");

		commentsWriter.writeStartElement("text");
		commentsWriter.writeStartElement("r");
		commentsWriter.writeStartElement("t");
		if (comment.charAt(0) < ' '
				|| comment.charAt(comment.length() - 1) == ' ')
			commentsWriter.writeAttribute("xml:space", "preserve");
		commentsWriter.writeCharacters(comment);
		commentsWriter.writeEndElement();// end t
		commentsWriter.writeEndElement();// end r
		commentsWriter.writeEndElement();// end text
		commentsWriter.writeEndElement();// end comment

		// write vml drawings
		vmlWriter.writeStartElement("v:shape");
		vmlWriter.writeAttribute("type", this.shapeTypeRefId);
		vmlWriter.writeAttribute("fillcolor", "#ffffe1");
		vmlWriter.writeAttribute("o:insetmode", "auto");

		vmlWriter.writeStartElement("x:ClientData");
		vmlWriter.writeAttribute("ObjectType", "Note");
		vmlWriter.writeStartElement("x:Row");
		vmlWriter.writeCharacters(String.valueOf(r));
		vmlWriter.writeEndElement();// end x:row
		vmlWriter.writeStartElement("x:Column");
		vmlWriter.writeCharacters(String.valueOf(c));
		vmlWriter.writeEndElement();// end x:column
		// <x:Anchor>
		// 1, 15, 0, 15, 3, 18, 3, 15</x:Anchor>
		vmlWriter.writeStartElement("x:Anchor");
		String anchorPoints = StringUtils.join(new Object[] {
				// start point(x,y)
				c + 1, 15, r, 15,
				// end point(x,y)
				c + 1 + 2, 15,
				r + 2 + new StringTokenizer(comment, "\n").countTokens(), 15 },
				",");
		vmlWriter.writeCharacters(anchorPoints);
		vmlWriter.writeEndElement();// end x:Anchor
		vmlWriter.writeEndElement();// end x:ClientData
		vmlWriter.writeEndElement();// end v:shape
	}
}
