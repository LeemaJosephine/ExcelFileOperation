package readAndWriteExcel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {
		
		// Get into the workbook
		XSSFWorkbook book = new XSSFWorkbook();
		
		// Create the sheet
		
		XSSFSheet sheet = book.createSheet("login");
		
		// Store the student details   -> Name(String) Age(int) City(String)
		
		Object[][] data = {    // 2-D Array
				
				{"Name","Age","City"},   // 1 array
				{"Ajay",20,"Delhi"},
				{"Arjun",25,"Chennai"},
				{"Anbu",23,"Mumbai"}	  // 4 rows and 3 columns
				
		};
		
		// Put the data into the sheet
		
		int rowCount = 0;
		
		// for each to get into each row
		
		for(Object[] row1 : data) {
			
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount=0;
			
			// for each to get the columns
			
			for(Object col:row1) {
				
				XSSFCell cell = row.createCell(columnCount++);
				
				// Checking the type of data and storing accordingly
				if(col instanceof String) {
					cell.setCellValue((String)col);
				}else if (col instanceof Integer) {
					cell.setCellValue((Integer) col);
				}
			}
			
		}
		
		try {
			FileOutputStream output = new FileOutputStream("C:\\Users\\Digital Suppliers\\eclipse-workspace\\ExcelFileOperation\\src\\main\\java\\readAndWriteExcel\\StudentDetails.xlsx");
			book.write(output);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		book.close();
	}

}
