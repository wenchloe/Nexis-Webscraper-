/*
 * 
 * NexisWebscraper takes in a company name/search query, a timeline to collect articles from, 
 * and a list of desired publication types. Opens a FireFox browser (can edit to be 
 * ChromeDriver if desired), logs in to the Nexis Uni database using University of 
 * Washington net-id, and searches the database for the search query. Filters based on 
 * the publication type, loops through for each publication type (may contain duplicates
 * which can be eliminated in Excel). Extracts title, date, publisher, word count, and 
 * all of the article's text. Inputs data into given excel workbook and sheet; prints
 * to the sheet after every page, or every ten articles collected. Ignores articles 
 * outside of the timeline. 
 * 
 * */

/*
 * See ReadMe for configuration details / list of packages and jars needed to be added 
 * */

// import Util and SimpleDateFormat (for Scanners + filtering irrelevant articles)
import java.util.*;
import java.util.concurrent.TimeUnit;
import java.text.SimpleDateFormat; 
// import Selenium WebDriver packages 
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
// import Apache POI Excel Workbook Packages (Excel 97-2003)
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.*; 


public class NexisWebscraper {

	public static final String UW_NETID_LOGIN = "https://weblogin.washington.edu/";
	public static final String NEXIS_UNI_URL = "log-in url"; // insert log-in link to database
	public static final String NETID = "net-id"; // insert net-id 
	public static final String PWD = "password"; // insert net-id password 
	public static final String SEARCH_QUERY = "INSERT KEYWORD HERE";
	public static final String[] PUBLICATION_TYPES = {"Newspapers", "Web-Based Publications", "Industry Trade Press",
			"Newswires & Press Releases", "Magazines & Journals", "News Transcripts",
			"Newsletters", "Aggregate News Sources"}; 
	private static final int NUM_COLUMNS = 6; // the number of columns for the Excel spreadsheet
	public static int rowNum = 1; // keeps track of current row number 
	private static final int yearFounded = 2012; // the cut-off year for collected articles
	private static final int monthFounded = 3; // the cut-off month for collected articles 

	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.gecko.driver", "c://path"); // insert path geckodriver here 
		DesiredCapabilities capabilities = DesiredCapabilities.firefox();
		capabilities.setCapability("marionette", true);

		// Opens Workbook and desired sheet 
		FileInputStream fis = new FileInputStream(new File("c://path")); // insert path to excel spreadsheet here
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		Sheet sheet = wb.getSheet("Sheet Name"); // insert sheet name here (must already be created)

		try {
			WebDriver driver = new FirefoxDriver();

			// opens up FireFox browser and logs in 
			driver.get(UW_NETID_LOGIN);
			driver.findElement(By.id("weblogin_netid")).sendKeys(NETID);
			System.out.println("logging in");
			driver.findElement(By.id("weblogin_password")).sendKeys(PWD);
			driver.findElement(By.name("submit")).click();

			// open up the Nexis Uni main page
			driver.get(NEXIS_UNI_URL);

			// wait for 15 seconds for the Nexis Uni page to load, then enter a query and search
			driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			driver.findElement(By.className("academic-searchterms")).sendKeys(SEARCH_QUERY);
			driver.findElement(By.className("BISGO")).click();

			// filter the results so that only English articles come up
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			driver.findElement(By.id("podfiltersbuttonlanguage")).click(); 
			WebElement elem = driver.findElement(By.xpath("//input[@data-value='English']"));
			JavascriptExecutor executor = (JavascriptExecutor) driver;
			executor.executeScript("arguments[0].click();", elem);
			
			// Collects articles by specified publication type 
			String prevPubType = "";
			for (int i = 0; i < PUBLICATION_TYPES.length; i++) {
				System.out.println("Going into for loop for: " + PUBLICATION_TYPES[i]);

				// wait for the page to load after choosing the English filter
				WebDriverWait wait = new WebDriverWait(driver, 40);
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@title='Language: English']")));

				System.out.println("Checking if the previous publication type filter needs to be cancelled");
				// cancel the previous publication type filter if this isn't the first iteration
				if (!prevPubType.equals("")) {
					System.out.println("Cancelling previous pub type filter");
					WebElement remove = driver.findElement(By.xpath("//button[@title='Publication Type: " + prevPubType + "']"));
					remove.click();
				}

				System.out.println("Opening the drop-down menu for publication type");
				// choose the new publication type
				// open up drop-down menu for publication type, then wait for page to load
				WebElement dropDownPubType = driver.findElement(By.id("podfiltersbuttonpublicationtype"));
				wait.until(ExpectedConditions.elementToBeClickable(dropDownPubType));
				dropDownPubType.click();

				System.out.println("Checking if the More button exists");
				// click "More" and open up more publication type options in the drop down if there are more options
				java.util.List<WebElement> elements = driver.findElements(By.xpath("//button[@data-action='moreless']"));
				if (elements.size() > 0) {
					System.out.println("Opened up more options in drop down");
					executor.executeScript("arguments[0].click();", elements.get(0));
				}

				// extract the data for the new publication type only if it can be found 
				// in the drop down menu; else, go to the next publication type
				System.out.println("Checking if " + PUBLICATION_TYPES[i] + " is in the drop down");
				elements = driver.findElements(By.xpath("//input[@data-value='" + PUBLICATION_TYPES[i] + "']"));
				if (elements.size() > 0) {
					System.out.println("Found " + PUBLICATION_TYPES[i] + " in drop down");

					System.out.println("Choosing " + PUBLICATION_TYPES[i]);
					// choose the next publication type in PUBLICATION_TYPES
					prevPubType = PUBLICATION_TYPES[i];
					executor.executeScript("arguments[0].click();", elements.get(0));
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@title='Publication Type: " + PUBLICATION_TYPES[i] + "']")));

					System.out.println("Processing articles");
					driver.findElement(By.xpath("//a[@aria-label='Next']")).click();
					// extract data from every article on the 1st page of search results and print to Excel
					processArticles(driver, wait, sheet);
					printToExcel(wb);

					// process the remaining pages of search results
					int pageNum = 2;
					System.out.println("trying to go to next page");
					while (driver.findElements(By.xpath("//a[@aria-label='Next']")).size() > 0) { // check if a next page exists	
						driver.findElement(By.xpath("//a[@aria-label='Next']")).click();
						wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@title='Language: English']")));		
						wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@data-value=" + "'" + (pageNum - 1) + "']")));
						
						// process and print articles in the current page 
						processArticles(driver, wait, sheet);
						printToExcel(wb);
						pageNum++;
						
					}
				} else {
					System.out.println("Did not find " + PUBLICATION_TYPES[i] + " in drop down");
					prevPubType = "";
				}
				System.out.println("---------------END OF FOR LOOP ITERATION FOR " + PUBLICATION_TYPES[i] + "----------------");
			}
		} catch (Exception e) {
			System.out.println(e);
		}
	}

	// Processes each article on the current page. Extracts the publisher / journal, publication date, 
	// title / headline, word count, and the body of the article. 
	private static void processArticles(WebDriver driver, WebDriverWait wait, Sheet sheet) {
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@title='Language: English']")));
		java.util.List<WebElement> articles = driver.findElements(By.xpath("//a[@data-action='title']"));
		int numOfArticlesOnPage = articles.size();
		
		// Repeats processing procedure for each article on the page 
		for (int i = 0; i < numOfArticlesOnPage; i++) {
			articles = driver.findElements(By.xpath("//a[@data-action='title']"));	
			if (!articles.isEmpty() && articles.size() - 1 >= i) {
			    WebElement elem = articles.get(i);
				wait.until(ExpectedConditions.elementToBeClickable(elem));
				elem.click();
				
				driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
				java.util.List<WebElement> headingElements = driver.findElements(By.className("SS_DocumentInfo"));
				
				String journal = headingElements.get(0).getText();
				String date = headingElements.get(1).getText();
				String title = driver.findElement(By.id("SS_DocumentTitle")).getText();
				
				// Checks if the date of the article is in the given timeline 
				if (dateIsValid(date)) {
					// If date is valid, extract the body and word count of the article 
					String body = driver.findElement(By.className("SS_LeftAlign")).getText();
					String wordCount = extractWordCount(body);
					String articleBody = extractArticleBody(body);
					
					// Stores all the information that will be printed to the sheet 
					int articleArraySize = 0; 
					articleArraySize = (articleBody.length() / 32767) + 1; 
					String[] article = divideString(articleBody);
					String[] articleData = new String[articleArraySize + (NUM_COLUMNS -1)];
					articleData[0] = SEARCH_QUERY;
					articleData[1] = journal;
					articleData[2] = date;
					articleData[3] = title;
					articleData[4] = wordCount;
					// In the case that the character count of the article exceeds Excel's
					// .xls file cell max character count, prints the remaining article body  
					// to a different row. 
					for (int k = 0; k < articleArraySize; k++) {
						articleData[5 + k] = article[k];
					}
					
					// Prints the data to the row of an Excel sheet (not the output file)
					Row row = sheet.createRow(rowNum);
					for (int j = 0; j < articleData.length; j++) {
						Cell cell = row.createCell(j);
						cell.setCellValue(articleData[j]);
					}
					rowNum++; 
				} 
				
				// return to the page of search results
				driver.navigate().back();
			} else {
				System.out.println("article" + i + " did not click");
			}
		}
	}
	
	// Accepts the string form of the date as such ("December 14, 2017 Thursday"). 
	// Returns whether the date is in the desired timeline (after the month and year 
	// of the cut off). 
	private static boolean dateIsValid(String date) {
		Scanner dateData = new Scanner(date); // scans the String and extracts month / year
		Scanner dateDataCheck = new Scanner(date); // checks whether day is included 
		
		// Checks how many tokens are in the date String 
		int numTokens = 0;
		while (dateDataCheck.hasNext()) {
			numTokens++;
			dateDataCheck.next();
		}
		// Extracts the month of the article's publication
		String month = dateData.next();
		// Skips over the date of the article's publication if provided
		if (numTokens > 2) {
			dateData.next();	
		} 
		// Extracts the year of the article's publication  
		int year = Integer.parseInt(dateData.next().substring(0,4));
		
		// Checks whether the article's publication occurred before the cut-off month / year
		if (year < yearFounded ) {
			return false; 
		} else if (year == yearFounded) {
			try {
		         Date monthData = new SimpleDateFormat("MMM").parse(month);//put your month name here
		         Calendar cal = Calendar.getInstance();
		         cal.setTime(monthData);
		         int monthNumber = cal.get(Calendar.MONTH) + 1;
		         if (monthNumber > monthFounded) {
		        	 	return false; 
		         }
		      } catch (Exception e) {
		    	  System.out.println("Exception occurred while processing date");
		      }
		}
		return true; 
	}
	
	// Takes in the body of the article and extracts a word count 
	private static String extractWordCount(String body) {
		String[] lines = body.split("\n");
		for (int i = 0; i < lines.length; i++) {
			String line = lines[i];
			if (!line.equals("")) { // ignore blank lines
				if (line.startsWith("Length: ")) {
					return line.substring(8, line.length() - 6);
				}
			}
		}
		
		return "";
	}
	
	// Takes in the article body and extracts the text from the body of the article, 
	// returning it as a String in a single line instead of paragraphs. 
	private static String extractArticleBody(String body) {
		String[] lines = body.split("\n");
		String data = "";
		boolean bodyStarted = false;
		for (int i = 0; i < lines.length; i++) {
			String line = lines[i];
			if (!line.equals("")) { // if the line is not empty, collect data 
				if (bodyStarted) {
					if (line.equals("Classification")) { // "Classification" signals end of text body
						return data;
					} else { 
						if (!line.endsWith(" ")) {
							line += " "; // Collects the entire line, appends to previous line 
						}
						data += line;
					}
				} else if (line.equals("Body")) { // signals the start of the body of the article
					bodyStarted = true;
				}
			}
		}
		
		return data;
	}
	
	// Divides the article body (in a single String) into different indexes of an array
	// based on the amount of characters in it. Excel's max cell character count for 
	// .xls documents is 32767 characters, so this stores 32765 characters per cell
	// until all of the text from the input has been placed into the array. 
	private static String[] divideString(String input) {
		int resultSize = (input.length() / 32767) + 1; 
		String[] result = new String[resultSize];
		int[] numCharacters = new int[resultSize]; 
		
		// calculates the number of characters per cell 
		for (int i = 0; i < resultSize - 1; i++) {
			numCharacters[i] = 32766; 
		}
		numCharacters[resultSize - 1] = input.length() - (32766 * (resultSize - 1));
		
		// divides the string into the array based on calculated character size per index 
		for (int i = 0; i < resultSize; i++) {
			result[i] = input.substring(0 + (32766 * i), numCharacters[i] - 1 + (32766 * i));
		}
		return result;
	}
	
	// Prints the given article data to the Excel sheet 
	public static void printToExcel(Workbook wb) throws IOException {
		try {
			FileOutputStream output = new FileOutputStream(new File("c://path")); // insert excel file name or path
			wb.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		} 
	}
} 
