package File;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperation {
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook book = new XSSFWorkbook(); //creating a book
		XSSFSheet sheet = book.createSheet();	//creating a sheet
		
		Object[] [] data = {	 //creating a data
				
				{"Name","Age","City"},
				{"Raj","26","Erode"},
				{"Muthu","66","Erode"},
				{"Arun","20","Chennai"}
	};
		int rowCount=0; // initializing at row count
		
		for(Object[] row : data) { 
			
			XSSFRow createRow = sheet.createRow(rowCount++);
			
		int columnCount=0;  // initializing at Column count
		
		for(Object column: row) {                   
			
			XSSFCell cell = createRow.createCell(columnCount++);
			
			if(column instanceof String)
			{ 
				cell.setCellValue((String) column);
			}
			else if(column instanceof Integer)
			{
				cell.setCellValue((Integer) column);
			} 
			
			try(                                                 
					FileOutputStream output = new FileOutputStream("C:\\Users\\mrajk\\eclipse-workspace\\ExcelFileOperation\\src\\main\\java\\File\\FileWrite.xlsx");){
					book.write(output);           
				}

			}
		}
	}
}		
		
		
		