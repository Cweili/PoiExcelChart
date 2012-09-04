package cweili;



import cweili.hssf.Chart;
import cweili.ss.OOXMLChart;

/**
 * 
 * @author cweili
 * @version 2012-8-22 下午1:08:00
 * 
 */
public class AllChart {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// 标题项目
		String[] titles = new String[] { "项目一", "项目二", "项目三" };
		// 数据
		double[] values = new double[] { 50, 20, 30 };
		
		String[] fileNames = new String[] { "bar", "pie", "barpie" };
		
		for(String fileName : fileNames) {
			// Excel 2003 图表
			System.out.println("读取 Excel 97-2003 模板文件 tpl_" + fileName + ".xls");
			Chart.createChart(titles, values, "tpl_" + fileName + ".xls", "out_" + fileName + ".xls");
			System.out.println("输出 Excel 97-2003 图表文件 out_" + fileName + ".xls");
			System.out.println("-------------------------------------------------");
			
			// Excel 2007 图表
			System.out.println("读取 Excel 2007 模板文件 tpl_" + fileName + ".xlsx");
			OOXMLChart.createChart(titles, values, "tpl_" + fileName + ".xlsx", "out_" + fileName + ".xlsx");
			System.out.println("输出 Excel 2007 图表文件 out_" + fileName + ".xlsx");
			System.out.println("-------------------------------------------------");
		}
	}
}
