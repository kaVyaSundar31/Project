package org.sam;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SampleExcel {
	public static void main(String[] args) throws IOException {
		
		//1.Mention the excel sheet path
		File f = new File("C:\\Users\\kaviy\\eclipse-workspace\\InmakesSampleProject\\Excel\\SampleData.xlsx");
		
		//2.To read  the File
		FileInputStream s = new FileInputStream(f);
		
		//3.To read .xlsx file
		Workbook wb = new XSSFWorkbook(s);
		
		//4.Get Sheet from WorkBook
		Sheet mySheet = wb.getSheet("Data");
		
		//5.Get row from sheet
		Row particularRow = mySheet.getRow(2);
		
		//6.Get cell from row
		Cell particularCell = particularRow.getCell(1);
		System.out.println(particularCell);
		
		
	}

}



