package HowToTakeSS;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.SearchContext;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.google.common.io.Files;

public class DateFornat {

	private static TakesScreenshot driver;

	public static void main(String[] args) throws IOException, InterruptedException {

		Date d=new Date ();
		System.out.println(d);
		
		//Simple Format
		//Date Format JAVA class
		DateFormat d1 = new SimpleDateFormat("MM/dd-yy & HH-mm-ss");
		String date = d1.format(d);

		//Screenshot
		
		 File f=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
	        
	        Files.copy(f, new File("C:\\Users\\Shubham\\Desktop\\Automation Practice\\Automation Screen Shots\\Quickpaylogin"+date+".jpg"));
	        
	       System.out.println ("Screenshot is Taken Login Page");
		
	
	} 

}
