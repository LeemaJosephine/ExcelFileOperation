package readAndWriteExcel;

import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		
		// Opening the book
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\Digital Suppliers\\eclipse-workspace\\ExcelFileOperation\\Data\\DataFile.xlsx");

		// Get to the sheet
		
		XSSFSheet sheet = book.getSheetAt(0);
		
		// get the no.of rows
		
		int rowCount = sheet.getLastRowNum();
		
		// get the no.of columns
		
		int columnCount = sheet.getRow(0).getLastCellNum();
		
		// iterate and get the cell value
		
		String[][] data = new String[rowCount][columnCount];
		
		for(int i=1;i<=rowCount;i++) {  // i = 1 when you want ignore the heading
			
			XSSFRow row = sheet.getRow(i);
			
			// get into columns
			
			for(int j=0;j<columnCount;j++) {
				
				XSSFCell cell = row.getCell(j);
				
				// read/get the value
				
				System.out.println(cell.getStringCellValue());
				
				// to store in array
				
				data[i-1][j] = cell.getStringCellValue();  // i = 1-1 = 0 j=0
				
			}
			System.out.println();
		}
		
//		for(String[] row : data) {
//			
//			for(String x : row) {
//				System.out.println(x+" ");
//			}
//			
//		}
		book.close();
	}

}
