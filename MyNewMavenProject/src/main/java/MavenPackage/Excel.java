package MavenPackage;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	XSSFSheet sh;
	
	public Excel() throws IOException {
		FileInputStream in=new FileInputStream("C:\\Users\\abdul\\OneDrive\\Documents\\Book1.xlsx");
		XSSFWorkbook w=new XSSFWorkbook(in);
		sh=w.getSheet("sheet1");
	}

	public String readData(int i, int j) {
		
		XSSFRow r=sh.getRow(i);
		XSSFCell c=r.getCell(j);
		int cellType=c.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_NUMERIC:{
			double d=c.getNumericCellValue();
			return String.valueOf(d);
		}
		
		case Cell.CELL_TYPE_STRING:{
			return c.getStringCellValue();
		}
				
		}
		
		return null;
	}
	
	

}
