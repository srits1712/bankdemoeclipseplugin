package basic;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

public class AllBrowsers {

	WebDriver driver;

		
		public void ff(){
		System.setProperty("webdriver.gecko.driver", "E:\\bindu network\\Library\\geckodriver.exe");
		driver=new FirefoxDriver();
		driver.get("http://www.google.com");
		System.out.println("opened in ff");
		}
		public void cd()
		{
		System.setProperty("webdriver.chrome.driver","E:\\bindu network\\Library\\chromedriver.exe");
		driver=new ChromeDriver();
		driver.get("http://www.google.com");
		System.out.println("opened in chrome");
		
		}
		
		public void ie()
		{
		System.setProperty("webdriver.ie.driver","E:\\bindu network\\Library\\IEDriverServer.exe");
		driver=new InternetExplorerDriver();
		driver.get("http://www.google.com");
		System.out.println("opened in ie");
		
		}
		public static void main(String[] args){
			
			AllBrowsers a=new AllBrowsers();
			a.ff();
			a.ie();
		}
	}


