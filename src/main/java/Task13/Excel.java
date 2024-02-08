
package Task13;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Excel {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet1 = book.createSheet();
		
		Object[][] data = {
				{"Name","Age","Email"},
				{"John Doe",30,"john@test.com"},
				{"Jane",28,"Jane@test.com"},
				{"bob smith",63,"jacky@example.com"},
				{"Swaonil",37,"Swapnil@example.com"},
		};
		
		int rowCount =0;
		
		for (Object[] row1 : data) {
			XSSFRow row = sheet1.createRow(rowCount++);	
			int columnCounr=0;
			
			for (Object col: row1) {
				
				XSSFCell cell = row.createCell(columnCounr++);
				
				if(col instanceof String) {
					cell.setCellValue((String) col);
				}else if (col instanceof Integer) {
					cell.setCellValue((Integer)col);
				}
				
			}
		}
		try(
			FileOutputStream output = new FileOutputStream("FIrstFile.xlsx");){
			book.write(output);
		}
		
	}

}
		
