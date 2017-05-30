package Test;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelTest {

	public static void main(String[] args) throws Exception {

		Workbook workbook = new HSSFWorkbook();

		Sheet first = workbook.createSheet("第一页");

		Row row = first.createRow(0);

		Cell cell = row.createCell(0);
		cell.setCellValue("张三");
		cell = row.createCell(1);
		cell.setCellValue("20");

		Sheet sec = workbook.createSheet("第2页");

		FileOutputStream fos = new FileOutputStream("C:/test.xls");
		workbook.write(fos);
		fos.close();
	}
}
