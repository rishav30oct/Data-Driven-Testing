package data_drivenTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Data_DrivenTest {
	
	
	@Test(dataProvider = "getTotalFlight")
	public void dataDriven(String src,String dest,int noOfFlight) throws Exception
	{
		//System.out.println("Departure=> "+src+" Arrival=> "+dest+" No of Flight=>"+noOfFlight);
		int date=10;
		String monthandyear="April 2022";
		WebDriverManager.chromedriver().setup();
		ChromeOptions option=new ChromeOptions();
		option.addArguments("--Disable-Notification");
		WebDriver driver=new ChromeDriver(option);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.get("https://www.makemytrip.com/");
		Actions action=new Actions(driver);
		action.moveByOffset(10, 10).click().perform();
		driver.findElement(By.xpath("//span[@class='langCardClose']")).click();
		driver.findElement(By.id("fromCity")).click();
		driver.findElement(By.xpath("//p[text()='"+src+", India']")).click();
		driver.findElement(By.xpath("//span[.='To']")).click();
		driver.findElement(By.xpath("//p[text()='"+dest+", India']")).click();
		driver.findElement(By.xpath("//div[@class='fsw_inputBox dates inactiveWidget ']")).click();
		driver.findElement(By.xpath("//div[text()='"+monthandyear+"']/ancestor::div[@class='DayPicker-Month']/descendant::p[text()='"+date+"']")).click();
		driver.findElement(By.xpath("//a[.='Search']")).click();
		List<WebElement> flightList = driver.findElements(By.xpath("//div[@class='makeFlex simpleow']"));
		int NoOfFlight=flightList.size();
		System.out.println("From =>"+src+" To =>"+dest+" TotalFlight =>"+NoOfFlight);
		FileInputStream fi=new FileInputStream("./src/test/resources/DataDriven.xlsx");
		Workbook wb1 = WorkbookFactory.create(fi);
		FileOutputStream fos=new FileOutputStream("./src/test/resources/DataDriven.xlsx");
		wb1.getSheet("data").getRow(noOfFlight).createCell(2).setCellValue(NoOfFlight+"");
		wb1.write(fos);
		wb1.close();
		fos.close();
		fi.close();
		
	}
	
	@DataProvider
	public Object[][] getTotalFlight() throws EncryptedDocumentException, IOException
	{
		FileInputStream fis=new FileInputStream("./src/test/resources/DataDriven.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		int lastRow = wb.getSheet("data").getLastRowNum();
		System.out.println(lastRow);
		
		int lastColomn=wb.getSheet("data").getRow(0).getLastCellNum();
		Object [][]objAry=new Object[lastRow][lastColomn];
		int k=0;
		int l=0;
		
		for(int i=1;i<=lastRow;i++)
		{
			String src=wb.getSheet("data").getRow(i).getCell(0).getStringCellValue();
			objAry[k++][0]=src;
		}
		for(int i=1;i<=lastRow;i++)
		{
			String dest=wb.getSheet("data").getRow(i).getCell(1).getStringCellValue();
			objAry[l++][1]=dest;
		}
		int m=0;
		for(int i=1;i<=lastRow;i++)
		{
			
			objAry[m++][2]=i;
		}
		return objAry;
	}
}
