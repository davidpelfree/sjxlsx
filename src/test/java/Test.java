import java.awt.Color;
import java.io.File;

import com.incesoft.tools.excel.ExcelRowIterator;
import com.incesoft.tools.excel.ReaderSupport;
import com.incesoft.tools.excel.WriterSupport;
import com.incesoft.tools.excel.support.CellFormat;
import com.incesoft.tools.excel.xlsx.ExcelUtils;

public class Test {

	public static void main(String[] args) {

		// check is office2007 or 03 version
		ExcelUtils.getExcelExtensionName(new File("/12.xlsx"));

		ReaderSupport rxs = ReaderSupport.newInstance(ReaderSupport.TYPE_XLSX, new File("/in.xlsx"));
		rxs.open();
		ExcelRowIterator it = rxs.rowIterator();
		while (it.nextRow()) {
			System.out.println(it.getCellValue(0));
		}
		rxs.close();

		WriterSupport wxs = WriterSupport.newInstance(WriterSupport.TYPE_XLSX, new File("/out.xlsx"));
		// WriterSupport wxs = WriterSupport.newInstance(WriterSupport.TYPE_XLS,
		// new File("/out.xls"));
		wxs.open();
		wxs.increaseRow();
		for (int i = 0; i < 5; i++) {
			wxs.increaseRow();
			wxs.writeRow(new String[] { "floydd" + i }, new CellFormat[] { new CellFormat(
					(i % 2 == 0 ? Color.PINK.getRGB() : Color.GREEN.getRGB()), -1, 0) });
		}
		wxs.close();
	}

}
