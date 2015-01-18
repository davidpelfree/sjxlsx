package com.incesoft.tools.excel.xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;

public class ExcelUtils {

	/**
	 * Excel 2007+ using the OOXML format(actually is a zip)
	 * 
	 * @return
	 */
	public static boolean isOOXML(InputStream inputStream) {
		try {
			return inputStream.read() == 0x50 && inputStream.read() == 0x4b && inputStream.read() == 0x03
					&& inputStream.read() == 0x04;
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * check excel version
	 * 
	 * @param file
	 * @return 'xlsx' for 07 or 'xls' for 03
	 */
	public static String getExcelExtensionName(File file) {
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(file);
			return isOOXML(stream) ? "xlsx" : "xls";
		} catch (IOException e) {
			throw new RuntimeException(e);
		} finally {
			if (stream != null) {
				IOUtils.closeQuietly(stream);
			}
		}
	}

	public static String checksumZipContent(File f) {
		ZipFile zipFile = null;
		try {
			zipFile = new ZipFile(f);
			Enumeration<? extends ZipEntry> e = zipFile.entries();
			List<Long> crcs = new ArrayList<Long>();
			while (e.hasMoreElements()) {
				ZipEntry entry = e.nextElement();
				crcs.add(entry.getCrc());
			}
			return DigestUtils.shaHex(StringUtils.join(crcs, ""));
		} catch (Exception e) {
			throw new RuntimeException("", e);
		} finally {
			try {
				if (zipFile != null)
					zipFile.close();
			} catch (IOException e) {}
		}
	}

	public static void main(String[] args) {
		File file = new File("/(全部-实例备份)(20120215154228)..xlsx");
		System.out.println(checksumZipContent(file));
		System.out.println(file.delete());
	}
}