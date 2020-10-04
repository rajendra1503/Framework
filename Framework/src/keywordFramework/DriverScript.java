package keywordFramework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
//import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.LogStatus;

public class DriverScript
{
	WebDriver driver;
	File file;
	FileInputStream inp;
	
	String fileExtn = null;
	String filePath = null;
	
	String sKeyword = null;
	String sLocator = null;
	String sLocatorValue = null;
	
	String execFlag = null;
	
	String tcSheetName = null;
	String fileName = null;
	
	Workbook wb = null;
	Sheet sht1 = null;
	Sheet sht2 = null;
	
	ExtentReports logger;
	
	public DriverScript(String path, String fileName) throws Exception
	{				
		this.fileName = fileName;
		
		filePath = path + "\\" + fileName;
		System.out.println(filePath);
		
		file = new File(filePath);
		
		try
		{
			inp = new FileInputStream(file);
		}
		catch(FileNotFoundException e)
		{			
			e.printStackTrace();
		}
		
		fileExtn = fileName.substring(fileName.indexOf("."));
		System.out.println(fileExtn);								
		
		if (fileExtn.contentEquals(".xlsx"))
		{
			wb = new XSSFWorkbook(inp);
		}
		else
		{
			wb = new HSSFWorkbook(inp);
		}
		
		logger = ExtentReports.get(DriverScript.class);
		logger.init("D:\\E Drive Backup\\WebDriver\\Screenshots\\first_report.html", true);
	}
	
	public void getExecutionFlag(String sheet1) throws Exception
	{				
		sht1 = wb.getSheet(sheet1);
		System.out.println(sheet1);
		int rows = sht1.getLastRowNum() - sht1.getFirstRowNum();
		System.out.println(rows);
		
		int cols = sht1.getRow(0).getLastCellNum() - sht1.getRow(0).getFirstCellNum();
		System.out.println(cols);
		
		for (int i = 1; i <= rows; i++)
		{
			Row row = sht1.getRow(i);
			execFlag = row.getCell(2).getStringCellValue();									
			
			if (execFlag.contentEquals("Y"))
			{
				tcSheetName = row.getCell(3).getStringCellValue();
				System.out.println(execFlag + "  " + tcSheetName);
				logger.startTest(tcSheetName);
				executeKeyword(tcSheetName);
				
			}
		}
	}
	
	public void executeKeyword(String sheetName)
	{
		sht2 = wb.getSheet(sheetName);
		
		int rows = sht2.getLastRowNum() - sht2.getFirstRowNum();
		System.out.println("No. of rows = " + rows);
		int cols = sht2.getRow(0).getLastCellNum() - sht2.getRow(0).getFirstCellNum();
		System.out.println("Column count = " + cols);
		
		for (int i = 1; i <= rows; i++)
		{
			Row row = sht2.getRow(i);
			sKeyword = row.getCell(2).getStringCellValue();
			System.out.println(sKeyword);
			sLocator = row.getCell(3).getStringCellValue();
			System.out.println(sLocator);
			sLocatorValue = row.getCell(4).getStringCellValue();
			System.out.println(sLocatorValue);
			
			switch (sKeyword)
			{
			case "openBrowser":
				String baseUrl = row.getCell(5).getStringCellValue();
				System.setProperty("webdriver.chrome.driver", "D:\\SeleniumSoft\\chromedriver_83\\chromedriver.exe");
				driver = new ChromeDriver();				
				driver.manage().window().maximize();
				driver.get(baseUrl);
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;
				
			case "setText":
				By textBox = getElementLocator(sLocator, sLocatorValue);
				String val = row.getCell(5).getStringCellValue();
				driver.findElement(textBox).sendKeys(val);
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;
			
			case "clickBtn":
				By cmdButton = getElementLocator(sLocator, sLocatorValue);
				driver.findElement(cmdButton).click();
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;
				
			case "wait":
				try {
					Thread.sleep(10000);
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				break;
				
			case "selectValue":
				By select = getElementLocator(sLocator, sLocatorValue);
				String choice = row.getCell(5).getStringCellValue();
				Select obj = new Select(driver.findElement(select));
				obj.selectByVisibleText(choice);
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;
				
			case "getPageTitle":
				String currentPageTitle = driver.getTitle();
				System.out.println(currentPageTitle);
				String eTitle = row.getCell(5).getStringCellValue();
				if (currentPageTitle.contentEquals(eTitle))
				{
					logger.log(LogStatus.PASS, row.getCell(1).getStringCellValue());
				}
				else
				{
					logger.log(LogStatus.FAIL, row.getCell(1).getStringCellValue());
				}				
				break;
				
			case "clickLink":
				By link = getElementLocator(sLocator, sLocatorValue);
				driver.findElement(link).click();
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;
				
			case "closeApplication":
				driver.quit();
				logger.log(LogStatus.INFO, row.getCell(1).getStringCellValue());
				break;				

			default:
				System.out.println("invalid keyword");
				break;
			}						
		}
		logger.endTest();
	}
	
	public By getElementLocator(String sLocator, String sLocatorValue)
	{
		switch (sLocator)
		{
			case "id":
				return By.id(sLocatorValue);
				
			case "name":
				return By.name(sLocatorValue);
				
			case "linkText":
				return By.linkText(sLocatorValue);
				
			case "partialLinkText":
				return By.partialLinkText(sLocatorValue);
				
			case "tagName":
				return By.tagName(sLocatorValue);
				
			case "cssSelector":
				return By.cssSelector(sLocatorValue);
				
			case "className":
				return By.className(sLocatorValue);
				
			case "xPath":
				return By.xpath(sLocatorValue);							
		}
		
		return null;
	}

	public static void main(String[] args) throws Exception
	{
		DriverScript obj = new DriverScript("D:\\E Drive Backup\\WebDriver\\TestData", "testcases.xlsx");
		obj.getExecutionFlag("TestSuite");		
		System.out.println("Done......");
	}

}
