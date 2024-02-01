package Test.Test1;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class migration {

	public static void main(String[] args) throws InterruptedException, IOException {
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://tn.unitetools.in");
		String excel = "C:\\Users\\sastry\\Desktop\\Sundry1.xlsx";
		String sname = "Sheet1";
		driver.findElement(By.xpath("//input[@id='user']")).sendKeys("TN12101048_de@coopsindia.com");
		driver.findElement(By.xpath("//input[@id='pwd']")).sendKeys("Unite@123");
		driver.findElement(By.xpath("//button[@id='btnvalidatelogin']")).click();
		Thread.sleep(2000);
		System.out.println("Test");
		driver.findElement(By.xpath("//div[@class='nav-subhead']/div/ul/li[2]")).click();
		driver.findElement(By.xpath("//div[@class='nav-subhead']/div/ul/li[2]/ul/li[3]")).click();
		driver.findElement(By.xpath("//div[@class='nav-subhead']/div/ul/li[2]/ul/li[3]/ul/li")).click();
     		
		FileInputStream inputStream = new FileInputStream(excel);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sname);
        
        int totalrow = sheet.getPhysicalNumberOfRows();
        for(int rowNum=1; rowNum<totalrow;rowNum++)
        {
        
    		Thread.sleep(2000);
    		driver.findElement(By.xpath("//div[@class='widget']/div[1]/div[1]/div/div/ul/li[2]")).click();
    		driver.findElement(By.xpath("//div[@class='widget']/div[1]/div[1]/div/div/ul/li[2]/ins")).click();
    		Thread.sleep(2000);
    		driver.findElement(By.xpath("//div[@class='widget']/div[1]/div[1]/div/div/ul/li[2]/ul/li[8]")).click();
    		WebElement el = driver.findElement(By.xpath("//div[@class='widget']/div[1]/div[1]/div/div/ul/li[2]/ul/li[8]/ins"));
    		el.click();
    		//el.sendKeys(Keys.DOWN);

    		//WebElement Element = driver.findElement(By.xpath("//div[@class='col-sm-12']//a[text()='115007000000000-Sundry Debtors for others(Annexure-XVI)']"));
    		
    		WebElement button = driver.findElement(By.id("node_114"));
    		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", button);
    		button.click();
    		Actions aa = new Actions(driver);
    		aa.contextClick(button).perform();
    		Thread.sleep(1000);
    		driver.findElement(By.xpath("//div[@id='contextMenu']/table/tbody/tr[1]/td/a")).click();
    		Thread.sleep(2000);
    	
    		Select dd = new Select(driver.findElement(By.xpath("//select[@id='Item2_IsLedger']")));
    		dd.selectByVisibleText("Account");
        	Row row = sheet.getRow(rowNum);
        	 String Aname = row.getCell(0).getStringCellValue();
             WebElement a = driver.findElement(By.xpath("//input[@id='Item2_LedgerName']"));
             a.click();
             a.sendKeys(Aname);
             org.apache.poi.ss.usermodel.Cell amt = row.getCell(1);
             String x ;
             if (amt != null && amt.getCellType() == CellType.NUMERIC) {
                 // If the cell is numeric, convert it to a string
                 x = String.valueOf(amt.getNumericCellValue());
             } else if (amt != null && amt.getCellType() == CellType.STRING) {
                 // If the cell is a string, get the string value
                 x = amt.getStringCellValue();
             } else {
                 // Handle other cell types as needed
                 x = ""; // or throw an exception, depending on your requirements
             }
             WebElement b = driver.findElement(By.xpath("//input[@id='Item2_OBDebitAmount']"));
             b.clear();
             b.sendKeys(x);
             workbook.close();
             inputStream.close();
             driver.findElement(By.xpath("//button[@value='Create Account']")).click();
             driver.findElement(By.xpath("(//div[@class='modal-content']//following-sibling::div)[26]/button")).click();

        }
       
	}

}
