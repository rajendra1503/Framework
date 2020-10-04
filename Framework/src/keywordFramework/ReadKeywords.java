package keywordFramework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class ReadKeywords
{
	WebDriver driver;	
	String sKeyword = null;
	String sLocator = null;
	String sLocatorValue = null;
	
	public void readExcel(String path, String fileName, String sheetName) throws Exception
	{
		driver = new FirefoxDriver();						
		
		String filePath = path + "\\" + fileName;
		
		File file = new File(filePath);
		FileInputStream inp = new FileInputStream(file);
		
		Workbook wBook = null;
		
		String fileExtn = fileName.substring(fileName.indexOf("."));				
		System.out.println("File extension is " + fileExtn);
		
		if (fileExtn.contentEquals(".xlsx"))
		{
			wBook = new XSSFWorkbook(inp);
		}
		else
		{
			wBook = new HSSFWorkbook(inp);
		}
		
		Sheet sht = wBook.getSheet(sheetName);
		
		int rows = sht.getLastRowNum() - sht.getFirstRowNum();
		System.out.println("No. of rows = " + rows);
		int cols = sht.getRow(0).getLastCellNum() - sht.getRow(0).getFirstCellNum();
		
		for (int i = 1; i <= rows; i++)
		{
			Row row = sht.getRow(i);
			sKeyword = row.getCell(2).getStringCellValue();
			sLocator = row.getCell(3).getStringCellValue();
			sLocatorValue = row.getCell(4).getStringCellValue();
			
			switch (sKeyword)
			{
			case "openBrowser":
				String baseUrl = row.getCell(5).getStringCellValue();
				driver.get(baseUrl);
				driver.manage().window().maximize();
				break;
				
			case "setText":
				By textBox = getElementLocator(sLocator, sLocatorValue);
				String val = row.getCell(5).getStringCellValue();
				driver.findElement(textBox).sendKeys(val);
				break;
			
			case "clickBtn":
				By cmdButton = getElementLocator(sLocator, sLocatorValue);
				driver.findElement(cmdButton).click();
				break;
				
			case "getPageTitle":
				String currentPageTitle = driver.getTitle();
				System.out.println(currentPageTitle);
				break;
				
			case "clickLink":
				By link = getElementLocator(sLocator, sLocatorValue);
				driver.findElement(link).click();
				break;
				
			case "closeApplication":
				driver.quit();
				break;				

			default:
				System.out.println("invalid keyword");
				break;
			}
			
			
		}
		
	}
	
	public By getElementLocator(String sLocator, String sLocatorValue)
	{
		if (sLocator.equals("id"))
		{
			return By.id(sLocatorValue);
		}
		else if(sLocator.equals("name"))
		{
			return By.name(sLocatorValue);
		}
		else if(sLocator.equals("linkText"))
		{
			return By.linkText(sLocatorValue);
		}
		else if(sLocator.equals("partialLinkText"))
		{
			return By.partialLinkText(sLocatorValue);
		}
		else if(sLocator.equals("tagName"))
		{
			return By.tagName(sLocatorValue);
		}
		else if(sLocator.equals("className"))
		{
			return By.className(sLocatorValue);
		}
		else if(sLocator.equals("cssSelector"))
		{
			return By.cssSelector(sLocatorValue);
		}
		else
		{
			return By.xpath(sLocatorValue);
		}
	}

	public static void main(String[] args) throws Exception
	{
		ReadKeywords obj = new ReadKeywords();
		obj.readExcel("E:\\WebDriver\\TestData", "testcases.xlsx", "Login Test case");
		obj.readExcel("E:\\WebDriver\\TestData", "testcases.xlsx", "WebOrderLogin");

	}

}
