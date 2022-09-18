package webscrapping;

import java.io.File;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;

public class WebScrapping_SuratSmartCity {

	static WebDriver driver;
	static String url = "http://office.suratsmartcity.com/SuratCOVID19/Home/COVID19BedAvailabilitydetails";
	static String zoneOption = "Central Zone"; // All Zones,West Zone,Central Zone,North Zone,East Zone - A,South Zone,
												// South West Zone,South East Zone,East Zone - B,South Zone-B

	// Excel variables
	static XSSFWorkbook suratCovidDataExcel;
	static XSSFSheet sheet;
	static XSSFRow row;
	static XSSFCell cell;
	static int rowNum = 0, cellNum;

	public static void main(String[] args) throws InterruptedException {
		setup();
		SelectZone();
		GetHospitalDetails();
		tearDown();
	}

	/****** Driver Initialization & Navigating to the url ******/
	public static void setup() {
		String driverPath = System.getProperty("user.dir") + "/drivers/chromedriver.exe";

		// Set the chrome driver path
		System.setProperty("webdriver.chrome.driver", driverPath);

		// Create instance of chrome driver
		driver = new ChromeDriver();

		// Invoke the url
		driver.get(url);
		driver.manage().window().maximize();

		CreateExcelWorkbook();
	}

	public static void tearDown() {
		driver.quit();
		WritetoExcel();
	}

	/******* Create Excel Workbook to export the web scrappping data to Excel *****/
	private static void CreateExcelWorkbook() {
		suratCovidDataExcel = new XSSFWorkbook();
		sheet = suratCovidDataExcel.createSheet("Surat Covid Details");

		cellNum = 0;

		// Create a header Row
		row = sheet.createRow(rowNum++);
		CreateExcelEntry("Zone");
		CreateExcelEntry("Hospital Name");
		CreateExcelEntry("Total Beds");
		CreateExcelEntry("Available Beds");
		CreateExcelEntry("Available O2 Beds");
		CreateExcelEntry("Available Ventilators");
		CreateExcelEntry("Contact Number");
		CreateExcelEntry("Contact Address");
		
		//Ignore the NUMBER_STORED_AS_TEXT error 
		sheet.addIgnoredErrors(new CellRangeAddress(0,9999,0,9999),IgnoredErrorType.NUMBER_STORED_AS_TEXT );

	}

	// Method for creating new cell and assigning the value
	private static void CreateExcelEntry(String cellValue) {
		row.createCell(cellNum++).setCellValue(cellValue);
	}

	// Write all the final data to excel
	private static void WritetoExcel() {
		String excelFilePath = System.getProperty("user.dir") + "/target/ScrappedData/SuratSmartCityWebScrappingData.xlsx";

		File outputFile = new File(excelFilePath);
		FileOutputStream fos;

		try {
			outputFile.createNewFile();
			fos = new FileOutputStream(outputFile);
			suratCovidDataExcel.write(fos); // write the data to excel
			suratCovidDataExcel.close();
			fos.flush();
			fos.close();

		} catch (Exception e) {
			e.printStackTrace();

		}
	}

	
	private static void SelectZone() {
		WebElement zoneDropDown = driver.findElement(By.id("ddlZone"));
		Select zones = new Select(zoneDropDown);
		zones.selectByVisibleText(zoneOption);
	}

	public static void GetHospitalDetails() throws InterruptedException {

		List<WebElement> hospitalList = driver.findElements(By.cssSelector(".card.custom-card"));
		String hospitalName, totalBeds, vacantBeds, contactNumber, contactAddress;
		String[] beds;

		Wait<WebDriver> wait = new FluentWait<WebDriver>(driver).withTimeout(Duration.ofSeconds(5))
				.pollingEvery(Duration.ofMillis(500));
		// .ignoring(NoSuchElementException.class);

		for (int i = 1; i <= hospitalList.size(); i++) {

			row = sheet.createRow(rowNum++); // Create a new row for every hospital in the list
			cellNum = 0; // Start the entry from 0th column (cell)

			System.out.println("\nZone : " + zoneOption);
			CreateExcelEntry(zoneOption);

			String parentLocator = "//div[@class='card custom-card'][" + i + "]";

			// When the contact modal pop up closes, allow time for main page controls to
			// load
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(parentLocator + "//a")));

			hospitalName = driver.findElement(By.xpath(parentLocator + "//a")).getText();

			System.out.println("Name of the hospital : " + hospitalName.replace("Contact", ""));
			CreateExcelEntry(hospitalName.replace("Contact", ""));

			/***** Search for additional hospital details relative to current hospital ****/

			// Total Beds
			totalBeds = driver.findElement(By.xpath(parentLocator + "//span[@class='count-text']")).getText();
			beds = totalBeds.split(" - ");
			System.out.println("Total Beds           :" + beds[1]);
			CreateExcelEntry(beds[1]);

			// Total Available Beds
			vacantBeds = driver.findElement(By.xpath(parentLocator + "//span[contains(@class,'pr-2')]")).getText();
			beds = vacantBeds.split(" - ");
			System.out.println("Total available Beds :" + beds[1]);
			CreateExcelEntry(beds[1]);

			// Open the collapsible panel with hospital bed details
			driver.findElement(By.xpath(parentLocator)).click();
			String collapsiblePanel = "//div[@id='collapseOne-" + i + "']";

			// Introduce a wait time for the collapsible panel to appear
			wait.until(ExpectedConditions.visibilityOfElementLocated(
					By.xpath(collapsiblePanel + "//div[text()='HDU(O2)']/following-sibling::div")));

			// O2 Beds
			WebElement o2Beds = driver
					.findElement(By.xpath(collapsiblePanel + "//div[text()='HDU(O2)']/following-sibling::div"));
			System.out.println("O2 Beds availibility :" + o2Beds.getText());
			CreateExcelEntry(o2Beds.getText());

			// Ventialtor(s) availability
			WebElement ventillators = driver
					.findElement(By.xpath(collapsiblePanel + "//div[text()='Ventilator']/following-sibling::div"));
			System.out.println("Ventialtor(s) availability :" + ventillators.getText());
			CreateExcelEntry(ventillators.getText());

			// Contact Details

			// Click on the hospital name to open the contact details pop up
			driver.findElement(By.xpath(parentLocator + "//a")).click();

			// Wait for the modal pop up to open before fetching value
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("modal-body")));

			contactNumber = driver.findElement(By.id("lblhosCno")).getText();
			contactAddress = driver.findElement(By.id("lblhosaddress")).getText();
			System.out.println("Contact Number     :" + contactNumber);
			System.out.println("Contact Address    :" + contactAddress);
			CreateExcelEntry(contactNumber);
			CreateExcelEntry(contactAddress);
			// Close the pop up
			driver.findElement(By.className("close")).click();

		}

	}

}
