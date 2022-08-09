package miniProject_873081;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxBinary;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.JavascriptExecutor;
public class Main {
	static WebDriver driver;
	//static WebDriver firefox_driver;
	public WebDriver createDriver() {
		//Function to create Chrome driver
		System.setProperty("webdriver.chrome.driver","C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\chromedriver.exe" );
        driver=new ChromeDriver();
        driver.manage().window().maximize(); //Maximizing the browser window
        driver.get("https://www.flipkart.com/"); //Opening the Flipkart website
        return driver;
	}
	public WebDriver createFirefoxDriver() {
		//Function to create Firefox driver
		System.setProperty("webdriver.gecko.driver","C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\geckodriver.exe" );
		FirefoxBinary firefoxBinary=new FirefoxBinary();
		//firefoxBinary.addCommandLineOptions("--headless");
		FirefoxProfile profile=new FirefoxProfile();
		FirefoxOptions firefoxOptions=new FirefoxOptions();
		firefoxOptions.setBinary(firefoxBinary);
		firefoxOptions.setProfile(profile);
		driver=new FirefoxDriver(firefoxOptions);
		driver.manage().window().maximize(); //Maximizing the browser window
		driver.get("https://www.flipkart.com/"); //Opening the Flipkart website
		return driver;
	}
	public WebDriver removePrompt() {
		//Function to remove login prompt
		Actions act=new Actions(driver);
        Action sendEsc=act.sendKeys(Keys.ESCAPE).build(); 
        sendEsc.perform(); //Removing the Login prompt
        return driver;
	}
	public WebDriver searchForMobiles() {
		//Function to search for mobiles through the Search textbox
		driver.findElement(By.name("q")).sendKeys("Mobiles"); //Typing 'Mobiles' in the Search textbox
        driver.findElement(By.className("vh79eN")).click(); //Clicking on Search button
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);//Waiting for page to load
        return driver;
	}
	public WebDriver selectPriceRequirement() {
		//Function to set max price as Rs.30,000
		WebElement Price=driver.findElement(By.xpath("//div[@class='_1YoBfV']//select[@class='fPjUPw']")); //This WebElement represents the max price dropdown in the Filters pane on the webpage
        Select maxPrice=new Select(Price);
        maxPrice.selectByValue("30000"); //Selecting max price as Rs.30000
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);//Wait for results to filter
        return driver;
	}
	public WebDriver sortByNewestFirst() {
		//Function to sort search results according to requirement
		driver.findElement(By.xpath("//div[contains(text(),'Newest First')]")).click(); //Clicking on Newest First sort option in the search results
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        return driver;
	}
	public WebDriver sortByNewestFirst1() {
		//Function to sort search results according to requirement
		WebElement element=driver.findElement(By.xpath("//div[contains(text(),'Newest First')]"));//.sendKeys(Keys.ENTER); //Clicking on Newest First sort option in the search results
		((JavascriptExecutor) driver).executeScript("return arguments[0].click();", element);
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        return driver;
	}
	public ArrayList<String> storeRawMobileDetails() {
		//Function to store (using an ArrayList<String>) all the mobile details (of 1st 5 mobiles) displayed in the search results (except price)
		//And close browser
		ArrayList<String> mobiles=new ArrayList<String>(); //ArrayList of String type to store details of mobiles
        String xpath_pre="//body/div[@id='container']/div/div[@class='t-0M7P _2doH3V']/div[@class='_3e7xtJ']/div[@class='_1HmYoV hCUpcT']/div[@class='_1HmYoV _35HD7C']/";
        String xpath_post="/div[1]/div[1]/div[1]/a[1]/div[2]/div[1]";
        //Getting details of 5 mobiles out of the search results
        for(int i=0;i<5;i++) {
        	String str="\n"+driver.findElement(By.xpath(xpath_pre+"div["+(i+2)+"]"+xpath_post)).getText();
        	mobiles.add(str);
        }
        System.out.println(mobiles); //Printing details of the 5 mobiles on console
        driver.close(); //Closing browser
        return mobiles;
	}
	public ArrayList<String[]> filterAndFormatMobileDetails(ArrayList<String> mobiles){
		//Function to filter relevant mobile details and format the input ArrayList<String> to ArrayList<String[]> which would be beneficial for organized insertion into excel file
		ArrayList<String[]> Mobile_details=new ArrayList<String[]>();
        //Editing mobile details and formatting
        Mobile_details.add(new String[]{"","Name","Storage","Display","Camera","Battery","Warranty"});
        for(int i=0;i<5;i++) {
        	String[] mob=mobiles.get(i).split("\n") ;
        	String[] mob_ed=new String[7];
        	int k=0;
        	for(int j=0;j<mob.length && k<7;j++) {
        		if(mob[j].contains("Ratings")) ;
        		else if(mob[j].contains("Processor"));
        		else if(mob[j].contains("Charger"));
        		else {
        			mob_ed[k]=mob[j];
        			k++;
        		}
        	}
        	Mobile_details.add(mob_ed);
        }
        return Mobile_details;
	}
	public void writeFileUsingPOI(ArrayList<String[]> data,String filepath) throws IOException {
		//Function to insert records/data into excel file
		//create blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); 	 
		//Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("Mobiles");	 	 
		//Iterate over data and write to sheet
		int rownum = 0;	 
		for (String[] mobiles : data){
			Row row = sheet.createRow(rownum++);
			int cellnum = 0;
			for (String mobile : mobiles){
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue(mobile);
			}
		}	  
		try{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File(filepath));
			workbook.write(out);
			out.close();
			System.out.println("Excel file has been created successfully");
		} catch (Exception e){
			e.printStackTrace();
		}finally {
			workbook.close();
		}
	}
	
	public void readAndDisplayExcelFile(String filepath) {
		//Function to read and display contents of Excel file
		try{
            FileInputStream file = new FileInputStream(new File(filepath)); 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file); 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0); 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();                 
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                            System.out.print(cell.getStringCellValue() + "\t");
                }
                System.out.println();
             }
             file.close();
             workbook.close();
        }            
        catch (Exception e) {
            e.printStackTrace();
        }
	}
	public static void main(String[] args) throws IOException{      
		Main m=new Main();
		m.createDriver();
		System.out.println("Chrome Driver created successfully. Navigation to flipKart homepage successful.");
		m.removePrompt();
		System.out.println("Login prompt removed.");
		m.searchForMobiles();
		System.out.println("Search for mobiles complete.");
		m.selectPriceRequirement();
		System.out.println("Max price set as Rs.30,000.");
		m.sortByNewestFirst();
		System.out.println("Search results sorted by 'Newest First'. \n");
		System.out.println("Details of the 5 mobile phones chosen:- ");
		ArrayList<String> mobiles=m.storeRawMobileDetails();
		System.out.println();
		System.out.println();
        ArrayList<String[]> Mobile_details=m.filterAndFormatMobileDetails(mobiles);
        //Function to write data to excel file
        m.writeFileUsingPOI(Mobile_details,"C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\MobileDetails_Chrome.xlsx");
        System.out.println();
        System.out.println();
        //Function to read data from excel file
        m.readAndDisplayExcelFile("C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\MobileDetails_Chrome.xlsx");
        
        
        m.createFirefoxDriver();
		System.out.println("Firefox Driver created successfully. Navigation to flipKart homepage successful.");
		m.removePrompt();
		System.out.println("Login prompt removed.");
		m.searchForMobiles();
		System.out.println("Search for mobiles complete.");
		m.selectPriceRequirement();
		System.out.println("Max price set as Rs.30,000.");
		m.sortByNewestFirst1();
		System.out.println("Search results sorted by 'Newest First'. \n");
		System.out.println("Details of the 5 mobile phones chosen:- ");
		ArrayList<String> mobiles1=m.storeRawMobileDetails();
		System.out.println();
		System.out.println();
        ArrayList<String[]> Mobile_details1=m.filterAndFormatMobileDetails(mobiles1);
        //Function to write data to excel file
        m.writeFileUsingPOI(Mobile_details1,"C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\MobileDetails_Firefox.xlsx");
        System.out.println();
        System.out.println();
        //Function to read data from excel file
        m.readAndDisplayExcelFile("C:\\Users\\PRAGYA\\Desktop\\RECRUITMENT\\cognizant\\Training\\Selenium\\Eclipse Workspace\\miniProject_873081\\MobileDetails_Firefox.xlsx");
    }
}

