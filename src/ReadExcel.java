
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
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
				
		//specify row and column used, get string value at location
		String employee1 = sheet1.getRow(2).getCell(0).getStringCellValue();
		
		//test name 1
		System.out.println("Employee name 1 " + employee1);
		
		//close wb, memory leak warning
		wb.close();
		//Something

	}

}
