package org.demo.demo2;


import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

public class mix1 {
	
	@Test
	public void main() throws InterruptedException, IOException, WriteException, BiffException
	{



	WebDriver driver= new FirefoxDriver();
		   driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
		driver.get("http://newtours.demoaut.com/");
		driver.manage().window().maximize();
		WritableWorkbook wwb= Workbook.createWorkbook(new File("demoaut.xls"));
		WritableSheet ws= wwb.createSheet("DataSheet", 0);
		wwb.getSheet(0);
		WebElement username= driver.findElement(By.name("userName"));
		
		Workbook wb1= Workbook.getWorkbook(new File("mahima.xls"));
		Sheet sheet1= wb1.getSheet(0);
		int rows= sheet1.getRows();
		
		for  (int i=0;i<rows;i++)
		{
			username.clear();
			username.sendKeys(sheet1.getCell(0,i).getContents());
			Thread.sleep(1000);
		}
	
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc1.png");
		WebElement password= driver.findElement(By.name("password"));
		password.clear();
		password.sendKeys("mercury");
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc2.png");
		WebElement login= driver.findElement(By.name("login"));
		login.click();
	
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc3.png");
		WebElement type= driver.findElement(By.xpath("//input[@value='oneway']"));
		type.click();
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc4.png");
		Select select = new Select(driver.findElement(By.name("passCount")) );
		select.selectByIndex(2);

		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc5.png");
		Select state=new Select( driver.findElement(By.name("fromPort")));
		List<WebElement> out= state.getOptions();
		int f= out.size();
		System.out.print("CITIES:"+"\n");
		int j=1;
		for (int i=0;i<f;i++)
		{
		
		System.out.print(out.get(i).getText()+"\n");
		Label label= new Label(0,j,out.get(i).getText());
		ws.addCell(label);
		j++;
		
		}
		Label label1 = new Label(0,0,"CITY");
		ws.addCell(label1);
		
	
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc6.png");
		WebElement service= driver.findElement(By.xpath("//input[@value='Business']"));
		service.click();
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc7.png");
		WebElement contnue= driver.findElement(By.xpath("/html/body/div/table/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/form/table/tbody/tr[14]/td/input"));
		contnue.click();
		WebElement depart= driver.findElement(By.xpath("//input[@value='Blue Skies Airlines$361$271$7:10']"));
		depart.click();
		WebElement continue2=driver.findElement(By.xpath("/html/body/div/table/tbody/tr/td[2]/table/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/form/p/input"));
		continue2.click();
		driver.findElement(By.name("passFirst0")).sendKeys("Mahima");
		driver.findElement(By.name("passLast0")).sendKeys("Aggarwal");
		Select meal= new Select( driver.findElement(By.name("pass.0.meal")));
		List<WebElement> meal1= meal.getOptions();
		int g=meal1.size();
		int k=1;
				for (int i=0;i<g;i++)
				{
					Label label= new Label(1,k,meal1.get(i).getText());
					ws.addCell(label);
					k++;
				
				}
				Label label2 = new Label(1,0,"MEAL PREFERENCE:");
				ws.addCell(label2);
				wwb.write();
				wwb.close();
				meal.selectByVisibleText("Hindu");
				driver.findElement(By.name("passFirst1")).sendKeys("Mahima");
				driver.findElement(By.name("passLast1")).sendKeys("Aggarwal");
				driver.findElement(By.name("passFirst2")).sendKeys("Mahima");
				driver.findElement(By.name("passLast2")).sendKeys("Aggarwal");
		driver.findElement(By.name("creditnumber")).sendKeys("000000000000000");
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc9.png");
		driver.findElement(By.name("buyFlights")).click();
		screenshot(driver,"C:/Users/Mahima Aggarwal/Desktop/selenium/screenshotabc10.png");
		Sheet sheet=wwb.getSheet(0);
		System.out.print("\n");
		for (int i=0;i<11;i++)
		{
			sheet.getCell(1, i);
			System.out.print(sheet.getCell(1, i).getContents()+"\n");
			
		}
	
		//driver.close();
		//driver.quit();
		
		}

	public static void screenshot(WebDriver driver, String path) throws IOException{
		
		TakesScreenshot camera = (TakesScreenshot) driver;
		File image= camera.getScreenshotAs(OutputType.FILE);
		FileUtils.copyFile(image,new File(path));
	}
}