package org.sam;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven2 {
	public static void main(String[] args) throws IOException {
		
		//1.Mention the excel sheet path
		File f = new File("C:\\Users\\kaviy\\eclipse-workspace\\InmakesSampleProject\\Excel\\SampleData.xlsx");
		
		//2.To read  the File
		FileInputStream s = new FileInputStream(f);
		
		//3.To read .xlsx file
		Workbook wb = new XSSFWorkbook(s);
		
		//4.Get Sheet from WorkBook
		Sheet mySheet = wb.getSheet("Sheet3");
		
		//To iterate all rows
		for (int i = 0; i < mySheet.getPhysicalNumberOfRows(); i++) {
			Row iterateRow = mySheet.getRow(i);
			
			//To iterate all cells
			for (int j = 0; j < iterateRow.getPhysicalNumberOfCells(); j++) {
				Cell iterateCell = iterateRow.getCell(j);
				int cellType = iterateCell.getCellType();
				//cellType = 1 --> string type
				//cellType = other than 1 ---> date cell/numeric cell
				
				if (cellType == 1) {
					String value = iterateCell.getStringCellValue();
						System.out.println(value);	
				}
				else if (DateUtil.isCellDateFormatted(iterateCell)) {
					Date d = iterateCell.getDateCellValue();
					SimpleDateFormat simp = new SimpleDateFormat("dd-MMM-yy");
					String format = simp.format(d);
					System.out.println(format);
				}
				else {
					double d = iterateCell.getNumericCellValue();
					long l = (long) d;
					String valueOf = String.valueOf(l);
					System.out.println(valueOf);
				}
				
			}
			
			
		}
		
	}

}

