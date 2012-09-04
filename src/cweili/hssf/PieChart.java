package cweili.hssf;

/**
 * 饼图
 * 
 * @author cweili
 * @version 2012-8-22 上午10:29:31
 * 
 */
public class PieChart extends Chart {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String inFile = "tpl_pie.xls"; // 模板文件
		String outFile = "out_pie.xls"; // 输出文件
		// 标题项目
		String[] titles = new String[] { "项目一", "项目二", "项目三" };
		// 数据
		double[] values = new double[] { 50, 20, 30 };
		System.out.println("读取模板 " + inFile);
		createChart(titles, values, inFile, outFile);
		System.out.println("输出文件 " + outFile);
	}
}
