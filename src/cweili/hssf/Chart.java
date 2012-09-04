package cweili.hssf;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 
 * @author cweili
 * @version 2012-8-22 上午11:30:13
 * 
 */
public class Chart {

	/**
	 * 修改模板数据并保存
	 * 
	 * @author cweili
	 * 
	 * @param titles
	 *            标题行
	 * @param values
	 *            数值
	 * @param inFile
	 *            模板文件
	 * @param outFile
	 *            输出文件
	 */
	public static void createChart(String[] titles, double[] values, String inFile,
			String outFile) {
		try {
			// 读取模板
			FileInputStream is = new FileInputStream(inFile);
			HSSFWorkbook wbs = new HSSFWorkbook(is);
			// 读取工作表0
			HSSFSheet sheet0 = wbs.getSheetAt(0);
			// System.out.println(sheet0.getPhysicalNumberOfRows());
			System.out.println("行数: " + (sheet0.getLastRowNum() + 1));
			// 标题项目
			HSSFRow titleRow = sheet0.getRow(0);
			for (int i = 0; i < titleRow.getLastCellNum(); ++i) {
				HSSFCell cell = titleRow.getCell(i);
				cell.setCellValue(titles[i]);
			}
			// 数据项目
			HSSFRow row = sheet0.getRow(1);
			// System.out.println(row.getPhysicalNumberOfCells());
			System.out.println("列数: " + row.getLastCellNum());
			for (int i = 0; i < row.getLastCellNum(); ++i) {
				HSSFCell cell = row.getCell(i);
				cell.setCellValue(values[i]);
			}
			// 输出文件
			FileOutputStream os = new FileOutputStream(outFile);
			wbs.write(os);
			is.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
