package DemoPackage;

import java.io.File;
import java.io.FileInputStream;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDate
{
	public static void main(String[] args)
	{
		try
		{
			FileInputStream file = new FileInputStream(new File("ExcelDates.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Reading date form cell "M10"
			// Row = 9
			// Column = M = 12

			// Row 1=0, 2=1, 3=2, 4=3, ...
			// Column A=1, B=2, C=3, ....

			Cell cell =  sheet.getRow(9).getCell(12);           
			System.out.println("Cell value: " + cell.getDateCellValue()); 

			//Creating calendar object from Date Object
			Calendar myCal = Calendar.getInstance();
			myCal.setTime(cell.getDateCellValue());
			System.out.println("Values from Calendar: " 
					+ myCal.get(Calendar.DAY_OF_MONTH) + "/" 
					+ myCal.get(Calendar.MONTH) + "/"
					+ myCal.get(Calendar.YEAR) );

			file.close();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
}