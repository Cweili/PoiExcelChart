package cweili.ss;

/**
 * OOXML 柱状图
 * 
 * @author cweili
 * @version 2012-8-22 上午11:56:56
 * 
 */
public class OOXMLBarChart extends OOXMLChart {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String inFile = "tpl_bar.xlsx"; // 模板文件
		String outFile = "out_bar.xlsx"; // 输出文件
		// 标题项目
		String[] titles = new String[] { "项目一", "项目二", "项目三" };
		// 数据
		double[] values = new double[] { 50, 20, 30 };
		System.out.println("读取模板 " + inFile);
		createChart(titles, values, inFile, outFile);
		System.out.println("输出文件 " + outFile);
	}
}
