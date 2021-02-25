package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class MasterBase {
	 public static WebDriver driver;

	// get driver
	public WebDriver getDriver() {

		System.setProperty("webdriver.chrome.driver", "D:\\Eclipse Workspace\\MavenProjecct\\Driver\\chromedriver.exe");
		driver=new ChromeDriver();
		return driver;
	}
	//LoadUrl
	public void loadurl(String url) {
		driver.get(url);
	}
	//BrowserMax
	public void browserMax() {
		driver.manage().window().maximize();
	}
	//SendKeys 	
	public void type(WebElement element,String data) {
		element.sendKeys(data);

	}
	//ClickButton
	public void click(WebElement element) {
		element.click();

	}
	//CLOSE
	public void close() {
			driver.close();
	}
	//Clear
	public void clear(WebElement element) {
		element.clear();
	}
	//excel Read	
	public String excelRead(String loc,String sheetname,int rowno,int colno) throws IOException {
		File file=new File(loc);
		FileInputStream stream=new FileInputStream(file);
		Workbook w=new XSSFWorkbook(stream);
		Sheet sheet = w.getSheet(sheetname);
		Row row = sheet.getRow(rowno);
		Cell cell = row.getCell(colno);
		int type = cell.getCellType();
		if (type==1) {
			String name = cell.getStringCellValue();
			return name;}

		if (type==0) {
			if (DateUtil.isCellDateFormatted(cell)) {


				String name = new SimpleDateFormat("dd-MM-yyyy").format(cell.getDateCellValue());
				return name;

			} else {
				String name = String.valueOf((long) cell.getNumericCellValue());
				return name;
			}
		}
		return null;
	}
	//EXCEL WRITE
	public void excelWrite(String location,String sheetname,int rowno,int colno,String data) throws IOException {
		File file=new File(location);
		Workbook w=new XSSFWorkbook();
		Sheet sheet = w.createSheet(sheetname);
		Row row = sheet.createRow(rowno);
		Cell cell = row.createCell(colno);
		cell.setCellValue(data);
		FileOutputStream stream=new FileOutputStream(file);
		w.write(stream);
		System.out.println("done...........");

	}
	//EXCEL UPDATE
	public void excelupdate(String location,String data,int rowno,int colno,String sheetname) throws IOException {
		File file=new File(location);
		FileInputStream stream=new FileInputStream(file);
		Workbook w=new XSSFWorkbook(stream);
		Sheet sheet = w.getSheet(sheetname);
		Row row = sheet.createRow(rowno);
		Cell cell = row.createCell(colno);
		cell.setCellValue(data);
		FileOutputStream stream2=new FileOutputStream(file);
		w.write(stream2);
		System.out.println("done...........");
			
		

	}














































}
