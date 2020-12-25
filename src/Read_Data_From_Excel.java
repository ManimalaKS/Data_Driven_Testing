import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Read_Data_From_Excel {

	public static void main(String[] args) throws IOException {
		
		String excelpath="C:\\Users\\Manimala\\Desktop\\Selenium\\Data_Driven_Testing\\Data\\countries.xlsx";
		FileInputStream fs=new FileInputStream(excelpath);
		XSSFWorkbook workbook=new XSSFWorkbook(fs);
		XSSFSheet sheet=workbook.getSheet("Sheet1");
		int rows=sheet.getLastRowNum();
		System.out.println(rows);
		int columns=sheet.getRow(1).getLastCellNum();
		System.out.println(columns);
		/*for(int r=0;r<=rows;r++)
		{
			XSSFRow row= sheet.getRow(r);
			for(int c=0;c<columns;c++)
			{
			XSSFCell cell= row.getCell(c);
			
			switch(cell.getCellType())
			{
			case STRING:
				System.out.print(cell.getStringCellValue());
				break;
			case NUMERIC:
				System.out.print(cell.getNumericCellValue());
				break;
			case BOOLEAN:
				System.out.print(cell.getBooleanCellValue());
				break;
			default:
				break;
			}
			System.out.print(" | ");
		
			}
			System.out.println();
		}
		*/
		// Using Iterator
		Iterator it=sheet.iterator();
		while(it.hasNext())
		{
		XSSFRow	row=(XSSFRow) it.next();
		Iterator celliterator=row.cellIterator();
		while(celliterator.hasNext())
		{
			XSSFCell cells=(XSSFCell) celliterator.next();
		    
			switch(cells.getCellType())
			{
			case STRING:
				System.out.print(cells.getStringCellValue());
				break;
			case NUMERIC:
				System.out.print(cells.getNumericCellValue());
				break;
			case BOOLEAN:
				System.out.print(cells.getBooleanCellValue());
				break;
			default:
				break;
			}
			System.out.print(" | ");
		
		}
		System.out.println();
		}
	}

}
