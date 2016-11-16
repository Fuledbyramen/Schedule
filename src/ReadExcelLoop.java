
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelLoop 
{

	public static void main(String[] args) throws Exception 
	{
		
		//location of excel file, import 'File' (java.io) class, specifying excel file sheet location
		File src = new File("C:\\Users\\w7260216\\Desktop\\Java Lab\\Schedule\\Schedule.xlsx");
		
		//specify file source, import 'FileInputStream' (java.io), add throw exception, base exception recommended
		FileInputStream fis = new FileInputStream(src);
		
		//create fileinputstream class object, coming from apache poi, changed to XSSFWorkbook on both sides, loads total workbook
		//XSSFWorkbook is needed for .xlsx files
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		//index 0 = sheet one, if needed, can get sheet 2 = getSheetAt(1), etc...
		XSSFSheet sheet1 = wb.getSheetAt(0);
				
		//check how many rows in sheet to run for loop
		int rowcount = sheet1.getLastRowNum();  //return number of rows you have in excel sheet, starts w/ zero
		
		//test
		System.out.println("Total rows is: " + rowcount);
		
		//for loop, starting at row 2 (row 0 is empty)
		for(int i = 2; i <= rowcount; i++){
			
			String data0 = sheet1.getRow(i).getCell(0).getStringCellValue();
			
			System.out.println("Test data from excel is " + data0);
		}
		
		
		
	}
}
