package Task13;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\RAVIVARMAN V\\Documents\\ExcelFileOperation\\FIrstFile.xlsx");
		XSSFSheet sheet1 = book.getSheetAt(0);
		
		int rowCount = sheet1.getLastRowNum();
		int columnCount = sheet1.getRow(0).getLastCellNum();
		
		Object[][] data= new String[rowCount][columnCount];
		
		
		
		for(int i=1;i<=rowCount;i++) {
			
			XSSFRow row = sheet1.getRow(i);
			
	
			
			for(int j=0;j<columnCount;j++) {
				
				XSSFCell cell = row.getCell(j);
				
				
				
				data[i-1][j] = cell.getStringCellValue(); 
				
				System.out.println(cell.getStringCellValue());
				
			}
			
		}
		
		book.close();
		
		
	}

}
