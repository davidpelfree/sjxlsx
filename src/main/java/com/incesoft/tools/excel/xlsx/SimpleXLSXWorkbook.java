package com.incesoft.tools.excel.xlsx;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.XMLStreamWriter;

import org.apache.commons.io.IOUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

/**
 * A simple implementation of OOXML(Excel part) to read and modify Excel 2007+
 * documents
 * 
 * @author floyd
 * 
 */
public class SimpleXLSXWorkbook {
	static {
		// this the fastest stax implementation by test,especially when doing
		// output
		if ("false".equals(System.getProperty("ince.tools.excel.disableXMLOptimize"))) {
			System.setProperty("javax.xml.stream.XMLInputFactory", "com.ctc.wstx.stax.WstxInputFactory");
			System.setProperty("javax.xml.stream.XMLOutputFactory", "com.ctc.wstx.stax.WstxOutputFactory");
		}
	}

	ZipFile zipfile;

	private InputStream findData(String name) {
		try {
			ZipEntry entry = zipfile.getEntry(name);
			if (entry != null) {
				return zipfile.getInputStream(entry);
			}
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return null;
	}

	private static final String PATH_XL_RELATION = "xl/_rels/workbook.xml.rels";

	private static final String PATH_XL_RELATION_SHEETS = "xl/worksheets/_rels/sheet%d.xml.rels";

	private static final String PATH_SHAREDSTRINGS = "xl/sharedStrings.xml";

	private static final String PATH_CONTENT_TYPES = "[Content_Types].xml";

	static private List<Pattern> blackListPatterns = new ArrayList<Pattern>();

	static private List<String> blackList = Arrays.asList(".*comments\\d+\\.xml", ".*calcChain\\.xml",
			".*drawings/vmlDrawing\\d+\\.vml");
	static {
		for (String pstr : blackList) {
			blackListPatterns.add(Pattern.compile(pstr));
		}
	}

	public SimpleXLSXWorkbook(File file) {
		try {
			this.zipfile = new ZipFile(file);
			InputStream stream = findData(PATH_SHAREDSTRINGS);
			if (stream != null) {
				parseSharedStrings(stream);
			}
			initSheets();
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	public void close() {
		if (this.zipfile != null)
			try {
				this.zipfile.close();
				this.zipfile = null;
			} catch (IOException e) {
				log.error("", e);
			}
		this.commiter = null;
		this.sharedStrings.clear();
		this.sharedStrings = null;
		this.fills.clear();
		this.fonts.clear();
		for (Sheet s : sheets) {
			s.cleanUp();
		}
		this.sheets.clear();
		this.styles.clear();
	}

	// Parse sharedStrings.xml >>>
	BidirectionMap sharedStrings = new BidirectionMap();

	// DualHashBidiMap sharedStrings = new DualHashBidiMap();

	int sharedStringLen = 0;

	private int addSharedString(String string) {
		if (string == null)
			throw new IllegalArgumentException("null string added to SharedStrings");

		Integer i = (Integer) sharedStrings.inverse().get(string);
		if (i != null) {
			return i;
		} else {
			sharedStrings.put(sharedStringLen++, string);
			return sharedStringLen - 1;
		}
	}

	/**
	 * sharedString -> rich text
	 * 
	 * @return
	 */
	public RichText createRichText() {
		RichText text = new RichText();
		int i = addRichText(text);
		text.setIndex(i);
		return text;
	}

	/**
	 * sharedString -> plain text
	 * 
	 * @param text
	 * @return
	 */
	public SharedStringText createPlainText(String text) {
		SharedStringText sharedStringText = new SharedStringText();
		sharedStringText.setIndex(addSharedString(text));
		sharedStringText.setText(text);
		return sharedStringText;
	}

	private int addRichText(RichText richText) {
		if (richText == null) {
			throw new IllegalArgumentException("null rich text added to sharedStrings");
		}
		sharedStrings.put(sharedStringLen++, richText);
		return sharedStringLen - 1;
	}

	String getSharedStringValue(int i) {
		Object obj = sharedStrings.get(i);
		if (obj instanceof String) {
			return (String) obj;
		}
		return null;
	}

	// int getSharedStringIndex(String string) {
	// Integer i = (Integer) sharedStrings.inverseBidiMap().get(string);
	// if (i != null) {
	// return i;
	// } else {
	// return -1;
	// }
	// }

	XMLInputFactory inputFactory = XMLInputFactory.newInstance();

	private String getSheetPath(int i) {
		return String.format(PATH_SHEET, i);
	}

	private String getSheetCommentPath(int i) {
		return String.format(PATH_SHEET_COMMENT, i);
	}

	private String getSheetCommentDrawingPath(int i) {
		return String.format(PATH_SHEET_COMMENT_VMLDRAWING, i);
	}

	private void parseSharedStrings(InputStream inputStream) throws Exception {
		XMLStreamReader reader = inputFactory.createXMLStreamReader(inputStream);
		int type;
		boolean si = false;
		StringBuilder builder = new StringBuilder(100);
		while (reader.hasNext()) {
			type = reader.next();
			switch (type) {
			case XMLStreamReader.CHARACTERS:
				if (si)
					builder.append(reader.getText());
				break;
			case XMLStreamReader.START_ELEMENT:
				if ("si".equals(reader.getLocalName())) {
					builder.setLength(0);
					si = true;
				}
				break;
			case XMLStreamReader.END_ELEMENT:
				if ("si".equals(reader.getLocalName())) {
					for (int j = builder.length() - 1; j >= 0; j--) {
						if (builder.charAt(j) == '_' && j - 6 >= 0 && builder.charAt(j - 5) == 'x'
								&& builder.charAt(j - 6) == '_') {
							builder.delete(j - 6, j + 1);
							j = j - 6;
						}
					}
					sharedStrings.put(sharedStringLen++, builder.toString());
					si = false;
				}
				break;
			default:
				break;
			}
		}
	}

	// Parse sharedStrings.xml <<<

	private static final String PATH_SHEET = "xl/worksheets/sheet%d.xml";

	private static final String PATH_SHEET_COMMENT = "xl/comments%d.xml";

	private static final String PATH_SHEET_COMMENT_VMLDRAWING = "xl/drawings/vmlDrawing%d.vml";

	private static final String PATH_STYLES = "xl/styles.xml";

	private static final String STR_XML_HEAD = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
	private static final byte[] DATA_XL_WORKSHEETS__RELS_SHEET = (STR_XML_HEAD + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>")
			.getBytes();

	XMLStreamReader getReader(String resourceId) {
		InputStream stream = findData(resourceId);
		if (stream == null) {
			if (resourceId.startsWith("xl/worksheets/_rels/sheet")) {
				byte[] b = new byte[DATA_XL_WORKSHEETS__RELS_SHEET.length];
				System.arraycopy(DATA_XL_WORKSHEETS__RELS_SHEET, 0, b, 0, b.length);
				stream = new ByteArrayInputStream(b);
			} else
				throw new RuntimeException("resource not found,resourceId=" + resourceId);
		}
		try {
			XMLStreamReader reader = inputFactory.createXMLStreamReader(stream);
			return reader;
		} catch (XMLStreamException e) {
			throw new RuntimeException(e);
		}
	}

	XMLStreamReader getSheetReader(Integer sheetId) {
		if (sheetId == null) {
			sheetId = 1;
		}
		return getReader(getSheetPath(sheetId));
	}

	XMLStreamReader getStylesReader() {
		return getReader(PATH_STYLES);
	}

	// SHEET>>>
	List<Sheet> sheets = new ArrayList<Sheet>();

	private void initSheets() {
		for (int i = 0; true; i++) {
			ZipEntry entry = zipfile.getEntry(getSheetPath(i + 1));
			if (entry == null) {
				break;
			}
			sheets.add(new Sheet(i, this));
		}
	}

	/**
	 * create new sheet added to exists sheet list
	 */
	public Sheet createSheet() {
		Sheet sheet = new Sheet(sheets.size(), this);
		sheets.add(sheet);
		return sheet;
	}

	public int getSheetCount() {
		return sheets.size();
	}

	/**
	 * Get sheet by index(0~sheetCount-1)
	 * 
	 * @param i
	 * @return
	 */
	public Sheet getSheet(int i) {
		return getSheet(i, true);
	}

	/**
	 * Get sheet by index(0~sheetCount-1)
	 * 
	 * @param i
	 * @param parseAllRow
	 *            true to load all rows;false for lazy loading without memory
	 *            consuming({@link Sheet#setAddToMemory(boolean)=false}) when
	 *            doing iterator by {@link Sheet#nextRow()}
	 * @return
	 */
	public Sheet getSheet(int i, boolean parseAllRow) {
		if (i >= sheets.size())
			throw new IllegalArgumentException("sheet " + i + " not exists!SheetCount=" + sheets.size());
		Sheet sheet = sheets.get(i);
		if (parseAllRow)
			sheet.parseAllRows();
		else
			sheet.setAddToMemory(false);
		return sheet;
	}

	// SHEET<<<

	// MODIFY >>>
	List<Font> fonts = new ArrayList<Font>();

	List<Fill> fills = new ArrayList<Fill>();

	List<CellStyle> styles = new ArrayList<CellStyle>();

	private static void writeStart(XMLStreamReader reader, XMLStreamWriter writer, String... attributes)
			throws XMLStreamException {
		writer.writeStartElement(reader.getLocalName());
		for (int i = 0; i < attributes.length; i += 2) {
			writer.writeAttribute(attributes[i], attributes[i + 1]);
		}

		if (reader != null) {
			for (int i = 0; i < reader.getNamespaceCount(); i++) {
				writer.writeNamespace(reader.getNamespacePrefix(i), reader.getNamespaceURI(i));
			}
			String attName;
			for (int i = 0; i < reader.getAttributeCount(); i++) {
				attName = reader.getAttributeLocalName(i);
				if (attributes.length > 0) {
					for (int j = 0; j < attributes.length; j += 2) {
						if (!attName.equals(attributes[j])) {
							if (reader.getAttributePrefix(i) != null
									&& reader.getAttributePrefix(i).length() > 0) {
								writer.writeAttribute(reader.getAttributePrefix(i), writer
										.getNamespaceContext().getNamespaceURI(reader.getAttributePrefix(i)),
										reader.getAttributeLocalName(i), reader.getAttributeValue(i));
							} else {
								writer.writeAttribute(reader.getAttributeLocalName(i),
										reader.getAttributeValue(i));
							}
						}
					}
				} else {
					if (reader.getAttributePrefix(i) != null && reader.getAttributePrefix(i).length() > 0) {
						writer.writeAttribute(reader.getAttributePrefix(i), writer.getNamespaceContext()
								.getNamespaceURI(reader.getAttributePrefix(i)), reader
								.getAttributeLocalName(i), reader.getAttributeValue(i));
					} else {
						writer.writeAttribute(reader.getAttributeLocalName(i), reader.getAttributeValue(i));
					}
				}
			}
		}
	}

	// private boolean addIndex(List list, IndexedObject obj, int indexOffset) {
	// if (obj == null)
	// return false;
	// int pos = list.indexOf(obj);
	// if (pos == -1) {
	// obj.setIndex(indexOffset + list.size());
	// list.add(obj);
	// return true;
	// } else {
	// obj.setIndex(((IndexedObject) list.get(pos)).getIndex());
	// return false;
	// }
	// }

	private static void writeDocumentStart(XMLStreamWriter writer) throws XMLStreamException {
		// PATCH ,cause XmlStreamRead's START_DOCUMENT event never occurred
		writer.writeStartDocument("UTF-8", "1.0");
	}

	/**
	 * <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	 * count="2" uniqueCount="2">
	 * 
	 * <si><t>啊上的发放</t></si> <si> <r><t>afa</t></r> <r><rPr><color
	 * rgb="FF0000"/></rPr><t>你好</t></r> <r><t>df</t></r></si>
	 * 
	 * </sst>
	 * 
	 * @param writer
	 * @throws XMLStreamException
	 */
	private void mergeSharedStrings(XMLStreamWriter writer) throws XMLStreamException {
		writeDocumentStart(writer);

		writer.writeStartElement("sst");
		writer.writeNamespace("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

		for (int i = 0; i < sharedStringLen; i++) {
			Object obj = sharedStrings.get(i);
			if (obj instanceof SharedStringText) {
				((SharedStringText) obj).serialize(writer);
			} else {
				try {
					writer.writeStartElement("si");
					writer.writeStartElement("t");
					writer.writeCharacters((String) obj);
					writer.writeEndElement();
					writer.writeEndElement();
				} catch (RuntimeException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

		writer.writeEndElement();// end sst
	}

	private void mergeStyles(XMLStreamWriter writer) throws XMLStreamException {
		prepareStylesCount();

		XMLStreamReader reader = getStylesReader();

		writeDocumentStart(writer);
		while (reader.hasNext()) {
			int event = reader.next();
			switch (event) {
			case XMLStreamReader.START_ELEMENT:
				// merge fonts , fills , styles
				if ("fonts".equals(reader.getLocalName())) {
					// Integer offset =
					// Integer.valueOf(reader.getAttributeValue(
					// null, "count"));
					// for (CellStyle cellStyle : styles) {
					// addIndex(fonts, cellStyle.getFont(), offset);
					// }
					writeStart(reader, writer, "count", String.valueOf(fontsCountOffset + fonts.size()));
				} else if ("fills".equals(reader.getLocalName())) {
					// Integer offset =
					// Integer.valueOf(reader.getAttributeValue(
					// null, "count"));
					// for (CellStyle cellStyle : styles) {
					// addIndex(fills, cellStyle.getFill(), offset);
					// }
					writeStart(reader, writer, "count", String.valueOf(fillsCountOffset + fills.size()));
				} else if ("cellXfs".equals(reader.getLocalName())) {
					// Integer offset =
					// Integer.valueOf(reader.getAttributeValue(
					// null, "count"));
					// for (CellStyle cellStyle : styles) {
					// addIndex(styles, cellStyle, offset);
					// }
					writeStart(reader, writer, "count", String.valueOf(stylesCountOffset + styles.size()));
				} else {
					writeStart(reader, writer);
				}
				break;
			case XMLStreamReader.END_ELEMENT:
				if ("fonts".equals(reader.getLocalName())) {
					for (Font font : fonts) {
						font.serialize(writer);
					}
				}
				if ("fills".equals(reader.getLocalName())) {
					for (Fill fill : fills) {
						fill.serialize(writer);
					}
				}
				if ("cellXfs".equals(reader.getLocalName())) {
					for (CellStyle cellStyle : styles) {
						cellStyle.serialize(writer);
					}
				}
				writer.writeEndElement();
				break;
			case XMLStreamReader.CHARACTERS:
				writer.writeCharacters(reader.getText());
				break;
			default:
				break;
			}
		}
	}

	private static final String NS_SHAREDSTRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

	private static final String NS_VMLDRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

	private static final String NS_COMMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

	private void mergeContentTypes(XMLStreamReader reader, XMLStreamWriter writer, String... overrides)
			throws XMLStreamException {
		// create sharedStrings relation if absent
		writeDocumentStart(writer);
		HashSet<String> existsTargets = new HashSet<String>();
		while (reader.hasNext()) {
			int type = reader.next();
			switch (type) {
			case XMLStreamReader.START_ELEMENT:
				if ("Override".equals(reader.getLocalName())) {
					existsTargets.add(reader.getAttributeValue(null, "PartName"));
				} else if ("Default".equals(reader.getLocalName())) {
					existsTargets.add(reader.getAttributeValue(null, "Extension"));
				}
				writeStart(reader, writer);
				break;
			case XMLStreamReader.END_ELEMENT:
				if ("Types".equals(reader.getLocalName())) {
					for (int i = 0; i < overrides.length; i += 3) {
						if (!existsTargets.contains(overrides[i + 1])) {
							writer.writeStartElement(overrides[i]);
							if ("Override".equals(overrides[i]))
								writer.writeAttribute("PartName", overrides[i + 1]);
							else
								writer.writeAttribute("Extension", overrides[i + 1]);
							writer.writeAttribute("ContentType", overrides[i + 2]);
							writer.writeEndElement();
						}
					}
				}
				writer.writeEndElement();
				break;
			case XMLStreamReader.START_DOCUMENT:
				writeStart(reader, writer);
				break;
			case XMLStreamReader.CHARACTERS:
				writer.writeCharacters(reader.getText());
				break;
			default:
				break;
			}
		}
		writer.close();
	}

	private static class RelationShipMerger {
		public String[] targets;// {target,type},...

		public RelationShipMerger(String[] targets) {
			super();
			this.targets = targets;
		}

		public Map<String, String> mergeRelation(XMLStreamReader reader, XMLStreamWriter writer)
				throws XMLStreamException {
			// create sharedStrings relation if absent
			writeDocumentStart(writer);
			int maxRid = 0;
			Integer rid;
			HashMap<String, String> existTargets = new HashMap<String, String>();
			while (reader.hasNext()) {
				int type = reader.next();
				switch (type) {
				case XMLStreamReader.START_ELEMENT:
					if ("Relationship".equals(reader.getLocalName())) {
						String ridStr = reader.getAttributeValue(null, "Id");
						rid = Integer.valueOf(ridStr.replace("rId", ""));
						if (rid > maxRid) {
							maxRid = rid;
						}
						existTargets.put(reader.getAttributeValue(null, "Target"), ridStr);
					}
					writeStart(reader, writer);
					break;
				case XMLStreamReader.END_ELEMENT:
					if ("Relationships".equals(reader.getLocalName())) {
						for (int i = 0; i < targets.length; i += 2) {
							if (!existTargets.containsKey(targets[i])) {
								// <Relationship Id="rId6"
								// Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
								// Target="sharedStrings.xml"/>
								writer.writeStartElement("Relationship");
								writer.writeAttribute("Type", targets[i + 1]);
								writer.writeAttribute("Target", targets[i]);
								writer.writeAttribute("Id", "rId" + String.valueOf(++maxRid));
								writer.writeEndElement();
								existTargets.put(targets[i], "rId" + maxRid);
							}
						}
					}
					writer.writeEndElement();
					break;
				case XMLStreamReader.START_DOCUMENT:
					writeStart(reader, writer);
					break;
				case XMLStreamReader.CHARACTERS:
					writer.writeCharacters(reader.getText());
					break;
				default:
					break;
				}
			}
			writer.close();
			return existTargets;
		}
	}

	int fontsCountOffset = 0;

	int fillsCountOffset = 0;

	int stylesCountOffset = 0;

	boolean stylesCountLoaded = false;

	private void prepareStylesCount() {
		if (stylesCountLoaded)
			return;
		stylesCountLoaded = true;

		try {
			XMLStreamReader reader = getStylesReader();
			loop1: while (reader.hasNext()) {
				int event = reader.next();
				switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if ("fonts".equals(reader.getLocalName())) {
						fontsCountOffset = Integer.valueOf(reader.getAttributeValue(null, "count"));
					}
					if ("fills".equals(reader.getLocalName())) {
						fillsCountOffset = Integer.valueOf(reader.getAttributeValue(null, "count"));
					}
					if ("cellXfs".equals(reader.getLocalName())) {
						stylesCountOffset = Integer.valueOf(reader.getAttributeValue(null, "count"));
						break loop1;
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					break;
				default:
					break;
				}
			}
		} catch (XMLStreamException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * Create gloabl font
	 * 
	 * @return
	 */
	public Font createFont() {
		return new Font();
	}

	public Fill createFill() {
		return new Fill();
	}

	public CellStyle createStyle(Font font, Fill fill) {
		if (font == null && fill == null) {
			throw new IllegalArgumentException("either font or fill is required");
		}
		if (font != null && font.getColor() == null) {
			throw new IllegalArgumentException("either font color required");
		}
		if (fill != null && fill.getFgColor() == null) {
			throw new IllegalArgumentException("either fill fgcolor required");
		}

		prepareStylesCount();

		int pos = -1;
		if (font != null) {
			pos = fonts.indexOf(font);
			if (pos == -1) {
				font.setIndex(fontsCountOffset + fonts.size());
				fonts.add(font);
			} else {
				font = fonts.get(pos);
			}
		}
		if (fill != null) {
			pos = fills.indexOf(fill);
			if (pos == -1) {
				fill.setIndex(fillsCountOffset + fills.size());
				fills.add(fill);
			} else {
				fill = fills.get(pos);
			}
		}

		CellStyle style = new CellStyle();
		style.setFill(fill);
		style.setFont(font);
		pos = styles.indexOf(style);
		if (pos == -1) {
			style.setIndex(stylesCountOffset + styles.size());
			styles.add(style);
		} else {
			style = styles.get(pos);
		}
		return style;
	}

	static class ModifyEntry {
		int r;

		int c;

		String comment;

		SharedStringText text;

		/**
		 * font for all string in the cell
		 */
		CellStyle style;

		public ModifyEntry(int r, int c, SharedStringText text, CellStyle style, String comment) {
			super();
			this.r = r;
			this.c = c;
			this.comment = comment;
			this.text = text;
			this.style = style;
		}

		public ModifyEntry(int r, int c, SharedStringText text, CellStyle style) {
			super();
			this.text = text;
			this.style = style;
			this.r = r;
			this.c = c;
		}
	}

	XMLOutputFactory outputFactory = XMLOutputFactory.newInstance();

	XMLStreamWriter newWriter(OutputStream outputStream) throws UnsupportedEncodingException,
			XMLStreamException {
		return outputFactory.createXMLStreamWriter(new OutputStreamWriter(outputStream, "UTF-8"));
	}

	abstract static class XMLStreamCreator {
		protected XMLStreamWriter writer;

		protected XMLStreamReader reader;

		public XMLStreamCreator() {
			super();
		}

		public XMLStreamReader createReader() throws XMLStreamException {
			return reader = createReaderInternal();
		}

		abstract XMLStreamReader createReaderInternal();

		abstract XMLStreamWriter createWriterInternal();

		public XMLStreamWriter createWriter() throws XMLStreamException {
			return writer = createWriterInternal();
		}

		public XMLStreamReader getReader() {
			return reader;
		}

		public XMLStreamWriter getWriter() {
			return writer;
		}
	}

	public static class Commiter {
		SimpleXLSXWorkbook wb;

		HashSet<String> mergedItems = new HashSet<String>();

		private Commiter(SimpleXLSXWorkbook wb, OutputStream output) {
			this.wb = wb;
			zos = new ZipOutputStream(output);
		}

		ZipOutputStream zos;

		boolean modified = false;

		public void beginCommit() {
		}

		XMLStreamWriter sheetWriter;

		Sheet lastCommitSheet;

		private class SheetCommentXMLStreamCreator extends XMLStreamCreator {
			ByteArrayOutputStream output;

			XMLStreamReader createReaderInternal() {
				throw new RuntimeException("not implemented");
			}

			XMLStreamWriter createWriterInternal() {
				output = new ByteArrayOutputStream();
				try {
					return writer = wb.outputFactory.createXMLStreamWriter(output);
				} catch (XMLStreamException e) {
					throw new RuntimeException("create xml stream writer failed", e);
				}
			}

			public ByteArrayOutputStream getOutput() {
				return output;
			}
		}

		SheetCommentXMLStreamCreator commentStreamCreator;

		SheetCommentXMLStreamCreator vmlStreamCreator;

		/**
		 * begin to write the sheet
		 * 
		 * @param sheet
		 * @throws IOException
		 * @throws XMLStreamException
		 */
		public void beginCommitSheet(Sheet sheet) throws IOException, XMLStreamException {
			sheetWriter = newWriter(wb.getSheetPath(sheet.getSheetIndex() + 1));
			lastCommitSheet = sheet;
			commentStreamCreator = new SheetCommentXMLStreamCreator();
			vmlStreamCreator = new SheetCommentXMLStreamCreator();
			sheet.writeSheetStart(sheetWriter, commentStreamCreator, vmlStreamCreator);
		}

		/**
		 * merge the sheet's modifications
		 */
		public void commitSheetModifications() {
			try {
				if (lastCommitSheet == null)
					throw new IllegalStateException("plz call beginCommitSheet(Sheet) first");
				lastCommitSheet.mergeSheet();
			} catch (XMLStreamException e) {
				throw new RuntimeException(e);
			}
		}

		/**
		 * Can be called more than once to write incremental rows. If the
		 * current sheet is modified,the modifications will be lost.So call
		 * {@link #commitSheetModifications()} before this method if u have any
		 * modifications.
		 */
		public void commitSheetWrites() {
			try {
				if (lastCommitSheet == null)
					throw new IllegalStateException("plz call beginCommitSheet(Sheet) first");
				lastCommitSheet.writeSheet();
			} catch (XMLStreamException e) {
				throw new RuntimeException(e);
			}
		}

		public void endCommitSheet() throws IOException, XMLStreamException {
			if (lastCommitSheet == null)
				throw new IllegalStateException("plz call beginCommitSheet(Sheet) first");
			int sheet = lastCommitSheet.getSheetIndex() + 1;
			String commentRid = null;
			String vmlRid = null;
			String settingRid = null;
			// merge relation
			String relationPath = String.format(PATH_XL_RELATION_SHEETS, sheet);
			ByteArrayOutputStream relation_baos = new ByteArrayOutputStream();
			String[] relationsToMerge = {};
			if (commentStreamCreator.getWriter() != null) {
				relationsToMerge = new String[] { "../comments" + sheet + ".xml", NS_COMMENT,
						"../drawings/vmlDrawing" + sheet + ".vml", NS_VMLDRAWING };
			}
			XMLStreamWriter writer = wb.newWriter(relation_baos);
			Map<String, String> rids = new RelationShipMerger(relationsToMerge).mergeRelation(
					wb.getReader(relationPath), writer);
			writer.close();
			relation_baos.close();
			commentRid = rids.get("../comments" + sheet + ".xml");
			vmlRid = rids.get("../drawings/vmlDrawing" + sheet + ".vml");
			settingRid = rids.get("../printerSettings/printerSettings" + sheet + ".bin");
			// end sheet commit
			lastCommitSheet.writeSheetEnd(commentRid, vmlRid, settingRid);
			sheetWriter.close();
			lastCommitSheet.clearRows();
			if (commentStreamCreator.getWriter() != null) {

				commentStreamCreator.getWriter().close();
				commentStreamCreator.getOutput().close();
				newZipEntry(wb.getSheetCommentPath(sheet));
				zos.write(commentStreamCreator.getOutput().toByteArray());

				vmlStreamCreator.getWriter().close();
				vmlStreamCreator.getOutput().close();
				newZipEntry(wb.getSheetCommentDrawingPath(sheet));
				zos.write(vmlStreamCreator.getOutput().toByteArray());
			}
			newZipEntry(relationPath);
			zos.write(relation_baos.toByteArray());
		}

		public void endCommit() {
			try {
				for (int i = 0; i < wb.sheets.size(); i++) {
					Sheet s = wb.sheets.get(i);
					if (s.isModified()) {
						modified = true;
						mergedItems.add(wb.getSheetPath(i + 1));
					}
				}
				if (modified) {
					commitStyles();
					mergeSheets();
					commitContentTypes();
					commitRelation();
					commitSharedStrings();
				}
				commitUnmodifiedStream();
				zos.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private void newZipEntry(String path) {
			try {
				zos.putNextEntry(new ZipEntry(path));
				mergedItems.add(path);
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private XMLStreamWriter newWriter(String path) {
			try {
				newZipEntry(path);
				return wb.newWriter(zos);
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		public void mergeSheets() {
			try {
				for (int i = 0; i < wb.sheets.size(); i++) {
					Sheet s = wb.sheets.get(i);
					if (s.isModified() && !s.isMerged()) {
						beginCommitSheet(s);
						s.mergeSheet();
						endCommitSheet();
					}
				}
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		// <Override PartName="/xl/sharedStrings.xml"
		// ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
		// <Override PartName="/xl/comments1.xml"
		// ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
		// <Default Extension="vml"
		// ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
		private void commitContentTypes() {
			try {
				XMLStreamWriter writer = newWriter(PATH_CONTENT_TYPES);
				ArrayList<String> parts = new ArrayList<String>();
				if (wb.sharedStringLen > 0) {
					parts.add("Override");
					parts.add("/xl/sharedStrings.xml");
					parts.add("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
				}
				boolean commentModified = false;
				for (int i = 0; i < wb.sheets.size(); i++) {
					Sheet s = wb.sheets.get(i);
					if (s.isCommentModified()) {
						commentModified = true;
						// comments
						parts.add("Override");
						parts.add("/xl/comments" + (s.getSheetIndex() + 1) + ".xml");
						parts.add("application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
						// vml drawings
					}
				}
				if (commentModified) {
					parts.add("Default");
					parts.add("vml");
					parts.add("application/vnd.openxmlformats-officedocument.vmlDrawing");
				}
				wb.mergeContentTypes(wb.getReader(PATH_CONTENT_TYPES), writer,
						parts.toArray(new String[parts.size()]));
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private void commitRelation() {
			try {
				XMLStreamWriter writer = newWriter(PATH_XL_RELATION);
				if (wb.sharedStringLen > 0)
					new RelationShipMerger(new String[] { "sharedStrings.xml", NS_SHAREDSTRINGS })
							.mergeRelation(wb.getReader(PATH_XL_RELATION), writer);
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private void commitStyles() {
			try {
				XMLStreamWriter writer = newWriter(PATH_STYLES);
				wb.mergeStyles(writer);
				writer.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private void commitSharedStrings() {
			try {
				XMLStreamWriter writer = newWriter(PATH_SHAREDSTRINGS);
				wb.mergeSharedStrings(writer);
				writer.close();
			} catch (Exception e) {
				throw new RuntimeException(e);
			}
		}

		private void commitUnmodifiedStream() throws IOException {
			Enumeration<? extends ZipEntry> entries = wb.zipfile.entries();
			loop1: for (; entries.hasMoreElements();) {
				ZipEntry entry = entries.nextElement();

				if (!modified || !mergedItems.contains(entry.getName())) {
					for (Pattern p : blackListPatterns) {
						if (p.matcher(entry.getName()).matches()) {
							continue loop1;
						}
					}
					zos.putNextEntry(new ZipEntry(entry.getName()));
					IOUtils.copy(wb.zipfile.getInputStream(entry), zos);
				}
			}
		}
	}

	Commiter commiter;

	public Commiter newCommiter(OutputStream output) {
		return commiter = new Commiter(this, output);
	}

	/**
	 * commit all modifications
	 * 
	 */
	public void commit(OutputStream output) throws Exception {
		if (commiter != null) {
			throw new IllegalStateException(
					"cannot commit again - newCommiter() or commit() has been called already.");
		}
		commiter = newCommiter(output);
		commiter.beginCommit();
		commiter.endCommit();
	}

	// MODIFY <<<

	// TEST
	public static void testMergeStyles(SimpleXLSXWorkbook excel, XMLStreamWriter writer) throws Exception {
		// CellStyle style = excel.createStyle();
		// style.setFont(new Font());
		// style.getFont().setColor("FFFF0000");
		// style = excel.createStyle();
		// style.setFont(new Font());
		// style.setFill(new Fill());
		// style.getFont().setColor("FF0000FF");
		// style.getFill().setFgColor("FF00FF00");

	}

	private static final Log log = LogFactory.getLog(SimpleXLSXWorkbook.class);

	@SuppressWarnings("unchecked")
	static private class BidirectionMap implements Map {
		private Map values = new LinkedHashMap();

		private Map inverseValues = new LinkedHashMap();

		public Map inverse() {
			return inverseValues;
		}

		public void clear() {
			values.clear();
			inverseValues.clear();
		}

		public boolean containsKey(Object key) {
			return false;
		}

		public boolean containsValue(Object value) {
			return false;
		}

		public Set entrySet() {
			throw new RuntimeException();
		}

		public Object get(Object key) {
			return values.get(key);
		}

		public boolean isEmpty() {
			return values.isEmpty();
		}

		public Set keySet() {
			throw new RuntimeException();
		}

		public Object put(Object key, Object value) {
			inverseValues.put(value, key);
			return values.put(key, value);
		}

		public void putAll(Map m) {
			throw new RuntimeException();
		}

		public Object remove(Object key) {
			throw new RuntimeException();
		}

		public int size() {
			return values.size();
		}

		public Collection values() {
			return values.values();
		}
	}

}
