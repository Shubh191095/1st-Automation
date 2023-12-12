package Snb;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class QuickPay {


	public static void main(String[] args) throws InterruptedException {
		
        WebDriver driver=new ChromeDriver();       
    	driver.manage().window().maximize();
		driver.get("http://10.180.184.30:8080/grs-web/#/login");
		Thread.sleep(2000);
	
		driver.findElement(By.id ("inputEmail3")).sendKeys("SITCHECKER2");
		Thread.sleep(2000);
		
		driver.findElement(By.name("password")).sendKeys("Ncb@1234");
		Thread.sleep(000);
		
		driver.findElement(By.xpath("//button[@class='btn btn-success btn-sm']")).click();
		Thread.sleep(5000);
		
        System.out.println("User login successful");
      
//Title compare Test case 1
        
        // String expectedTitle = "Saudi National Bank";
        		
        //String actualTitle=	driver.getTitle();
        
        	   // if(expectedTitle.equals(actualTitle	))
        	   // {
        	   //	System.out.println("actualTitle is matched with expectedTitle ");
        	   //  }
        
        // else
        	   //{
        	   //System.out.println("Test case is failed");
        	   // }
        
        //driver.close();
		
        //System.out.println("End Of Program");
        
		   //driver.close();
      
   // Test case 2 one time one run honar 1 or 2
        
        String expectedURL ="http://10.180.184.30:8080/grs-web/app.html#/changePass";
        String actualURL=driver.getCurrentUrl();
        
       if(expectedURL.equals(actualURL))
       {
    	   System.out.println("Test case is passedd");
    	   
       }
       else
       { 
    	   System.out.println("Test case is failed");
    	   
    	   driver.close();
   		 
   		   System.out.println("End Of Program");
           
           driver.close();
        
	}
}
}
