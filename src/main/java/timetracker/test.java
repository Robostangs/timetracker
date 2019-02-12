package timetracker;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;

public class test
{
	public static void main(String[] args) throws IOException
	{
		InputStream inp = new FileInputStream("workbook.xls");
		//InputStream inp = new FileInputStream("workbook.xlsx");
		 
		Workbook wb = WorkbookFactory.create(inp);
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(2);
		Cell cell = row.getCell(3);
		if (cell == null)
		    cell = row.createCell(3);
		cell.setCellValue("a test");
		 
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("workbook.xls");
		wb.write(fileOut);
		fileOut.close();
	

		
		
	}
	

}
