package timetracker;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.swing.JFileChooser;

import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.util.Date;
import java.util.Iterator;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.ArrayList;
public class ExcelWriter
{


	public static void main(String[] args) throws FileNotFoundException, IOException 
	{
		
		ArrayList<ArrayList<String> > table = new ArrayList<ArrayList<String>>();
		
		
		JFileChooser fileChooser = new JFileChooser();
		int returnValue = fileChooser.showOpenDialog(null);
		HSSFWorkbook workbook = null;
		try 
		{
			workbook = new HSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
		}
		 catch(Exception e) 
		{
			System.out.println("Please restart the program and choose a valid .xls file");
			
		}
			
		Sheet sheet = workbook.getSheetAt(0);
		//this is the code that gets the sheet, everything after this works fine except for the end...
		
		if(returnValue == JFileChooser.APPROVE_OPTION)
		{
			
			for(Iterator<Row> rit = sheet.rowIterator(); rit.hasNext();)
			{
				Row row = rit.next();
				ArrayList<String> list = new ArrayList<String>();
				for(Iterator<Cell> cit = row.cellIterator(); cit.hasNext(); )
				{
					DataFormatter dataFormatter = new DataFormatter();

					Cell cell = cit.next();

					
					list.add(dataFormatter.formatCellValue(cell));
				}

				table.add(list);

			}	
			System.out.println(table);
			
			
		}
		else
		{
			System.out.println("Please choose a valid XLS document");
		}

		


		System.out.println("(Type 'exit' to exit program)");
		while (true)
		{
			int column = column(table);
			column(table);
			String id = userInput();

			if(id.equals("exit"))
			{
				break;
			}
			
			
			boolean SignIn = false;
			System.out.println("Are you signing in or out?[in, out]");
			Scanner input = new Scanner(System.in);
			
			System.out.println("");
			System.out.println("Searching for your ID...");
			if(input.next().equals("in"))
			{
				SignIn = true;
			}
			if(SignIn == true)
			{
				int counter = 0;
				for(int i = 0; i < table.size(); i++)
				{

					if(id.equals(table.get(i).get(2)))
					{
						counter = 1;
						Row row = sheet.getRow(i);
						if(row.getPhysicalNumberOfCells()%2 == 0) 
						{

						
							Cell cell = row.createCell(column - 1);
							DateFormat df = new SimpleDateFormat("hh:mm aa");
							Date dateobj = new Date();
							cell.setCellValue(df.format(dateobj));
							System.out.println("Welcome, you have signed in!");
						}
						else
						{
							System.out.println("You have already signed in!");
						}




					}


				}
				if(counter == 0)
				{
					System.out.println("Your ID was not found in the spreadsheet");
				}
			}
			else
			{
				int counter = 0;
				for(int i = 0; i < table.size(); i++)
				{

					if(id.equals(table.get(i).get(2)))
					{
						counter = 1;


						Row row = sheet.getRow(i);
						if(row.getPhysicalNumberOfCells()%2 == 1) 
						{

							Cell cell = row.createCell(column);
							DateFormat df = new SimpleDateFormat("hh:mm aa");
							Date dateobj = new Date();
							cell.setCellValue(df.format(dateobj));
						}
						else
						{
							System.out.println("You haven't signed in yet!");
						}




					}

				}
				if(counter == 0)
				{
					System.out.println("Your ID was not found in the spreadsheet");
				}
			}
			
			
		}
		
		
		try {
			  
			  workbook.write(new FileOutputStream(fileChooser.getSelectedFile()));
			  workbook.close();
		    } catch(Exception e) 
		{
			e.printStackTrace();
		}
			}
	//this is the try that keeps throwing me an error every single time
	
	public static String userInput() 
	{
			Scanner input = new Scanner(System.in); 
			System.out.println("Please enter your ID number");
			String id = input.next();
			DateFormat df = new SimpleDateFormat("hh:mm aa");
		    Date dateobj = new Date();
		    System.out.println(df.format(dateobj));
		    System.out.println();
		    //add code for name not found error
		    
		    return id;
	}
	
	public static void autoFill(ArrayList<ArrayList<String> > table) {
		DateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
		Date dateobj1 = new Date();
		
		
	}
		
	public static int column(ArrayList<ArrayList<String> > table)
	{
		
		DateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
		Date dateobj1 = new Date();
		int column = 0;


		for(int i = 0; i<table.get(0).size(); i++)
		{
			if(String.valueOf(table.get(0).get(i)).equals(df1.format(dateobj1)) )
			{
			    column = i;
			    


			}
		}	
		return column;
		
	}
	
	
	
	
	public void reference()
	{
		//Makes workbook		
				Workbook workbook = new HSSFWorkbook();
			//Makes sheet in workbook, name goes in parenthesis 
			//use "WorkbookUtil.createSafeSheetName("  ")" To bypass any naming errors	
				Sheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("Robostangs attendence"));
			
			//creates row called row in the sheet as well as a cell called cell in the sheet
				Row row = sheet.createRow(1);
				Cell cell = row.createCell(4);
			//(cellname).setCellValue("  ");	
				cell.setCellValue("Phillip");
				
				
				System.out.println(cell.getRichStringCellValue().toString());
				
			
				
				
			//This code creates a new excel file in your project folder	
				try {
					FileOutputStream output = new FileOutputStream("Test14.xls");
					workbook.write(output);
					output.close();
					System.out.println("printed output successfully");
				    } catch(Exception e) 
				{
					e.printStackTrace();
				}
					}
				
}