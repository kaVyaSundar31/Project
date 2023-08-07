package org.sam;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;
	
	public static void launchBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}
	
	
	public static void windowMaximize() {
		driver.manage().window().maximize();
	}
	
	public static void launchUrl(String Url) {
		driver.get(Url);
	}
	
	public static String pageTitle() {
		String title = driver.getTitle();
		return title;
	}

	public static String pageUrl() {
		String Url = driver.getCurrentUrl();
		return Url;
	}

	public static void passText(String txt, WebElement ele) {
		ele.sendKeys(txt);	
	}
	
	public static void closeEntireBrowser() {
		driver.quit();
	}
	
	public static void clickBtn(WebElement ele) {
		ele.click();
	}
	
	public static void screenShot(String imgName) throws IOException {
		TakesScreenshot ts = (TakesScreenshot) driver;
		File img = ts.getScreenshotAs(OutputType.FILE);
		File f = new File("location+imgName.png");
		FileUtils.copyFile(img, f);
	}
	
	public static Actions a;
	
	public static void moveTheCursor(WebElement targetElement) {
		a = new Actions(driver);
		a.moveToElement(targetElement).perform();
	}
	
	public static void dragDrop(WebElement dragWebElement, WebElement dropWebElement) {
		a = new Actions(driver);
		a.dragAndDrop(dragWebElement, dropWebElement).perform();
	}
	
	public static JavascriptExecutor js;
	
	public static void scrollThePage(WebElement tarWebElement) {
		js = (JavascriptExecutor)driver;
		js.executeScript("argument[0].scrollIntoView(true)", tarWebElement);
	}

	public static void scroll(WebElement element) {
		js = (JavascriptExecutor)driver;
		js.executeScript("argument[0].scrollIntoView(false)", element);
	}
		
	public static void excelRead(String sheetName, int rowNum, int cellNum) throws IOException {
		File f = new File("excellocation.xlsx");
		FileInputStream s = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(s);
		Sheet mySheet = wb.getSheet("Data");
		Row r = mySheet.getRow(rowNum);
		Cell c = r.getCell(cellNum);	
		int cellType = c.getCellType();
		
		String value = " ";
		if (cellType == 1) {
			String value2 = c.getStringCellValue();
			
		}
		
		else if (DateUtil.isCellDateFormatted(c)) {
			Date dd = c.getDateCellValue();
			SimpleDateFormat simp = new SimpleDateFormat(value);
			String value1 = simp.format(dd);
		}
		
		else {
			double d = c.getNumericCellValue();
			long l = (long) d;
			String valueOf = String.valueOf(l);
		}
			
	}
	
	public static void createNewExcelFile(int rowNum, int cellNum, String writeData) throws IOException {
		File f = new File("C:\\Users\\kaviy\\eclipse-workspace\\InmakesSampleProject\\Excel\\NewFile.xlsx");
		Workbook wb = new XSSFWorkbook();
		Sheet newSheet = wb.createSheet("Datas");
		Row newRow = newSheet.createRow(rowNum);
		Cell newCell = newRow.createCell(cellNum);	
		newCell.setCellValue(writeData);
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
	}
	
	public static void createCell(int getRow, int creCell, String newData) throws IOException {
		File f = new File("C:\\\\Users\\\\kaviy\\\\eclipse-workspace\\\\InmakesSampleProject\\\\Excel\\\\NewFile.xlsx");
		FileInputStream s = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(s);
		Sheet mySheet = wb.getSheet("Datas");
		Row r = mySheet.getRow(getRow);
		Cell c = r.createCell(creCell);
		c.setCellValue(newData);
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
	}
	
	public static void createRow(int creRow, int creCell, String newData) throws IOException {
		File f = new File("C:\\\\Users\\\\kaviy\\\\eclipse-workspace\\\\InmakesSampleProject\\\\Excel\\\\NewFile.xlsx");
		FileInputStream s = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(s);
		Sheet mySheet = wb.getSheet("Datas");
		Row r = mySheet.createRow(creRow);
		Cell c = r.createCell(creCell);
		c.setCellValue(newData);
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
	}
	
	public static void updateDataToParticularCell(int getTheRow, int getTheCell,String exisitingData,  String writeNewData) throws IOException {
		File f = new File("C:\\Users\\kaviy\\eclipse-workspace\\InmakesSampleProject\\Excel\\NewFile.xlsx");
		FileInputStream s = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(s);
		Sheet mySheet = wb.getSheet("Datas");
		Row r = mySheet.getRow(getTheRow);
		Cell c = r.getCell(getTheCell);
		String str = c.getStringCellValue();
		if (str.equals(exisitingData)) {
			c.setCellValue(writeNewData);
		}
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
	}

		
}	

