package Test;

import java.io.FileInputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Excel内容的抽取
 * 
 * @author zhangjinfeng
 * @date 2017年5月25日下午8:02:45 TODO
 */
public class ExcelExtract {

	public static void main(String[] args) throws Exception {

		// 第一步获取要导入的Excel

		FileInputStream stream = new FileInputStream("C:\\Users\\Administrator\\Desktop\\2017年面试学生信息汇总.xls");

		// 第二部创建POIFSFileSystem

		POIFSFileSystem system = new POIFSFileSystem(stream);

		// 第三步利用POIFSFileSystem创建HSSFWorkBook

		HSSFWorkbook workbook = new HSSFWorkbook(system);

		// 创建Excel导出对象

		ExcelExtractor extractor = new ExcelExtractor(workbook);

		// 去除sheet的名字

		extractor.setIncludeSheetNames(false);
		System.out.println("root " + extractor.getRoot());

		System.out.println(extractor.getText());
	}
}
