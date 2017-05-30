package Test;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//哈希表之链地址法

public class Test {

	public static void main(String[] args) throws IOException {

		Workbook workbook = new HSSFWorkbook();

		Sheet sheet = workbook.createSheet("第一页");

		Row row = sheet.createRow(0);

		Cell cell = row.createCell(0);

		cell.setCellValue("张三");

		CellStyle style = workbook.createCellStyle();

		style.setFillBackgroundColor(IndexedColors.BLACK.getIndex());

		Font font = workbook.createFont();

		font.setColor(Font.COLOR_RED);
		style.setFont(font);

		cell.setCellStyle(style);

		FileOutputStream fileOutputStream = new FileOutputStream("C:/111.xls");

		workbook.write(fileOutputStream);
	}
}
