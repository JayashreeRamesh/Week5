package week5.testNG.Homeassignments;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public String[][] excelRead(String fileName) throws IOException {
		
		//Step 1 : Setup the excel document path 
		XSSFWorkbook wb = new XSSFWorkbook("./data/"+ fileName +".xlsx");
		
		//Step 2 : Setup the excel worksheet
		XSSFSheet ws = wb.getSheet("Sheet1");
		
		//Step 3 : To find the number of rows in the worksheet
		int rowCount = ws.getLastRowNum();
		
		//Step 4 : To find the number of cells in a row
		
		short cellCount = ws.getRow(1).getLastCellNum();
		
		//Declare 2D Array
		String[][] data = new String[rowCount][cellCount];
		
		for (int i = 1; i <=rowCount; i++) {
			
			for (int j = 0; j < cellCount; j++) {
				
				String text = ws.getRow(i).getCell(j).getStringCellValue();
				System.out.println(text);
				data[i-1][j] = text;
				
			}
			
		}
		
		wb.close();
		return data;
		
		
		

	}

}
