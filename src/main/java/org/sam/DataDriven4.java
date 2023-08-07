package org.sam;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven4 {
	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\kaviy\\eclipse-workspace\\InmakesSampleProject\\Excel\\newFile.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet newSheet = w.createSheet("Datas");
		Row newRow = newSheet.createRow(0);
		Cell newCell = newRow.createCell(0);
		newCell.setCellValue("DataDriven");
		FileOutputStream fos = new FileOutputStream(f);
		w.write(fos);
			
	}

}
