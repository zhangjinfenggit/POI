package Test;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * Excel 导入功能的实现
 * 
 * @author zhangjinfeng
 * @date 2017年5月25日下午8:01:21 TODO
 */
public class EntryExcel {

	public static void main(String[] args) throws Exception {

		// 第一步创建输入流
		FileInputStream fis = new FileInputStream("C:\\Users\\Administrator\\Desktop\\2017年面试学生信息汇总.xls");

		POIFSFileSystem fileSystem = new POIFSFileSystem(fis);

		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileSystem);

		HSSFSheet sheet = hssfWorkbook.getSheetAt(0);

		for (int i = 0; i < sheet.getLastRowNum(); i++) {

			HSSFRow row = sheet.getRow(i);

			for (int j = 0; j < row.getLastCellNum(); j++) {

				HSSFCell cell = row.getCell(j);
				System.out.print(getValue(cell) + "  ");
			}
			System.out.println();

		}

	}

	private static String getValue(HSSFCell cell) {

		if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			return cell.getNumericCellValue() + "";
		} else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {

		}
		return cell.getStringCellValue();
	}
}
