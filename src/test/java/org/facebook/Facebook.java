package org.facebook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Facebook {
    public static File file;
  	public static FileInputStream fin;
  	public static FileOutputStream fos;
  	public static Sheet sheet;
  	public static Row row;
  	public static Cell cell;

  	public static Cell c;
  	public static WebDriver driver;	
	public static Workbook w;

	
	public static void loadBrowser() {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
	}
	
	public static void launchUrl(String url) {
		driver.get(url);
	}
	
	public static void maximizePage() {
		driver.manage().window().maximize();

	}
	
	public static void btnClick(WebElement element2) {
		element2.click();
	}
	public static void sendData(WebElement element, String txt) {
	     element.sendKeys(txt);
	}
  //WAITS
    public static void pauseExecution(long milliseconds) {
 	try {
 		Thread.sleep(milliseconds);
 	} catch (InterruptedException e) {
 		// TODO Auto-generated catch block
 		e.printStackTrace();
 		}
 	}
    public static void applyWaitToAllElement() {
 	   driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
    }
  	/* Create Rows and Cells */
	public static void createRowsAndCells(String fileName, String SheetName, int rowNo, int cellNo, String value)
			throws IOException {
		File f = new File("C:\\Users\\Admin\\eclipse-workspace\\JunitPrograms\\Excel" + fileName + ".xlsx");

		
		
		if (rowNo == 0 && cellNo == 0) {
			w = new XSSFWorkbook();
			Sheet s = w.createSheet(SheetName);
			Row ro = s.createRow(rowNo);
			c = ro.createCell(cellNo);
		} else if (rowNo >= 0 && cellNo > 0) {
			fin = new FileInputStream(f);
			w = new XSSFWorkbook(fin);
			Sheet s = w.getSheet(SheetName);
			Row ro = s.getRow(rowNo);
			c = ro.createCell(cellNo);
		} else if (rowNo> 0 && cellNo == 0) {

			fin = new FileInputStream(f);
			w = new XSSFWorkbook(fin);
			Sheet s = w.getSheet(SheetName);
			Row ro = s.createRow(rowNo);
			c = ro.createCell(cellNo);
		}
		c.setCellValue(value);
		fos = new FileOutputStream(f);
		w.write(fos);
		System.out.println("RowandColumn Done");
	}
  	
}
