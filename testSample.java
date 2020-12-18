import java.io.IOException;
import java.util.ArrayList;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class testSample {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		dataDriven d = new dataDriven();
		ArrayList<String> data = d.getData("Add Profile");
		System.out.println(data.get(0));
		String key =data.get(1);
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		
		//Getting the data from excel and feeding into our test scenario 
		System.setProperty("webdriver.chrome.driver", "D:\\Personal_docs\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.google.com/");
		driver.findElement(By.xpath("//input[@title='Search']")).sendKeys(key);
	}

}
