package com.automation.app;

import java.io.File;
import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
/*import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;*/
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws InterruptedException, IOException
    {
    	System.setProperty("webdriver.gecko.driver","/home/kalana/Downloads/selenium-java-3.14.0/geckodriver");
    	
    	// Create a new instance of the Firefox driver
		WebDriver driver = new FirefoxDriver();
    	    	
    	File excelFile = new File("contacts.xlsx");
        FileInputStream fis = new FileInputStream(excelFile);

        // we create an XSSF Workbook object for our XLSX Excel File
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        

       for(int i = 0 ; i < workbook.getNumberOfSheets() ; i++) {
    	// we get first sheet
           XSSFSheet sheet = workbook.getSheetAt(i);
           
           // we iterate on rows
           Iterator<Row> rowIt = sheet.iterator();
           while(rowIt.hasNext()) {
        	   Row row = rowIt.next();
        	   Cell cell1,cell2,cell3 = null;
        	   cell1 = row.getCell(0);
               cell2 = row.getCell(1);
               cell3 = row.getCell(2);
               
              
               if(cell1.toString().equals("go_to_url")) {
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("element id:"+cell2.toString());
            	   //Launch the Online Store Website
           		 driver.get(cell2.toString());
           		// Print a Log In message to the screen
                System.out.println("Successfully opened the website "+cell2.toString());
               }else if(cell1.toString().equals("send_keys_input_by_name")){
            	   
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("element id:"+cell2.toString());
            	   System.out.println("element value:"+cell3.toString());
            	  
            	   WebElement element = driver.findElement(By.name(cell2.toString()));
             	   element.sendKeys(cell3.toString());
               }else if(cell1.toString().equals("button_click_by_text")){
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("button text:"+cell2.toString());  
            	   WebElement element = driver.findElement(By.xpath("//button[contains(text(),'"+isNumeric(cell2.toString())+"')]"));
                   element.click();
               }else if(cell1.toString().equals("input_click_by_name")){
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("input value:"+cell2.toString());  
            	   WebElement element = driver.findElement(By.xpath("//input[@value='"+isNumeric(cell2.toString())+"' and @type='submit']"));
                   element.click();
               }else if(cell1.toString().equals("link_click_by_text")){
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("input value:"+cell2.toString()); 
            	   
            	   WebElement element = driver.findElement(By.xpath("//a[contains(text(),'"+isNumeric(cell2.toString())+"')]"));
                   element.click();
               }else if(cell1.toString().equals("link_click_by_href_and_text")){
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("text:"+cell3.toString()); 
            	   System.out.println("href value:"+cell2.toString()); 
            	   
            	   WebElement element = driver.findElement(By.xpath("//a[contains(text(),'"+isNumeric(cell3.toString())+"') and @href='"+isNumeric(cell2.toString()+"']")));
                   element.click();
               }else if(cell1.toString().equals("send_keys_checkbox_by_name")){
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("element id:"+cell2.toString());  
            	   System.out.println("status:"+cell3.toString());  
            	   if(cell3.toString().equals("checked")) {
            		   System.out.println("status:"+cell3.toString());  
            		   WebElement element = driver.findElement(By.name(isNumeric(cell2.toString())));
                	   element.click(); 
            	   }
               }else if(cell1.toString().equals("send_keys_dropdown_by_name")) {
            	   System.out.println("command:"+cell1.toString());
            	   System.out.println("element id:"+cell2.toString());  
            	 
            	   int x = (int) cell3.getNumericCellValue();
            	   System.out.println("selected value:"+x);  
            	               	   
                   System.out.println("Select value"+ x);          

            	   WebElement dropdown = driver.findElement(By.name(isNumeric(cell2.toString()))); 
                   Select select = new Select(dropdown); 
                   java.util.List<WebElement> options = select.getOptions(); 
                   for(WebElement item:options) 
                   { 
                	   if((""+x) == item.getText()) {
                		   Select element = new Select(driver.findElement(By.name(isNumeric(cell2.toString())))); 
                    	   element.selectByVisibleText(""+x);
                	   }else {
                           System.out.println("Dropdown value: "+ item.getText());          
                	   }
                   }  
               }
        	   //System.out.println(cell1.toString()+" - "+cell2.toString()+";");
           }
       }
               
        fis.close();
    
        //Wait for 5 Sec
  		Thread.sleep(5000);
  		
        // Close the driver
        driver.quit();
    }
    
        
    public static String isNumeric(String str) {
        try {
            int number = (int) Float.parseFloat(str);
            System.out.println("isNumeric value:"+number+""); 
            return number+"";
        } catch (Exception e) {
        	System.out.println("isNumeric value:"+str); 
            return str;
        }
    }
    

}
