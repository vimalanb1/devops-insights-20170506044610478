package IBM;

import static org.junit.Assert.fail;

import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Properties;
import java.util.Set;
import java.util.StringTokenizer;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;
import javax.swing.JOptionPane;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.UnexpectedAlertBehaviour;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import org.xml.sax.SAXException;
//import org.apache.log4j.Logger;

public class Autorun{
	private static WebDriver driver;

	private static boolean tobecontinue = true;
	private static StringBuffer verificationErrors = new StringBuffer();
	private static FileInputStream fileInputStream;
	private static FileOutputStream outfile;

	public static String browsername;
	public static String TxtID = "";
	public static String browserpath = "src/browser.properties";
	public int numOfTestcases = 0;



	@Parameters("browser")
	@BeforeMethod
	// @Parameters("browser")
	public void setUp(String browser) throws Exception {
		Properties props = new Properties();
		props.load(new FileInputStream(browserpath));


		if (isProcessRunging("EXCEL.EXE")) {
			Runtime.getRuntime().exec("taskkill /IM EXCEL.EXE");
		}
		if (browser.equals("firefox")) {

			/*
			 * ProfilesIni allProfiles = new ProfilesIni(); FirefoxProfile
			 * profile = allProfiles.getProfile("Selenium");
			 */
			driver = new FirefoxDriver();

			fileInputStream = new FileInputStream(props.getProperty("firefox.excelPath"));
			browsername = "firefox";
		} else if (browser.equals("iexplorer")) {
			System.setProperty("webdriver.ie.driver", props.getProperty("iexplorer.webdriverPath"));

			DesiredCapabilities caps = DesiredCapabilities.internetExplorer();


			caps.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
			caps.setCapability(CapabilityType.UNEXPECTED_ALERT_BEHAVIOUR,UnexpectedAlertBehaviour.ACCEPT);




			driver = new InternetExplorerDriver(caps);

			fileInputStream = new FileInputStream(props.getProperty("iexplorer.excelPath"));
			browsername = "iexplorer";
		} else if (browser.equals("iexplorerprivate")) {
			System.setProperty("webdriver.ie.driver", props.getProperty("iexplorer.webdriverPath"));

			DesiredCapabilities caps = DesiredCapabilities.internetExplorer();
			caps.setCapability(InternetExplorerDriver.FORCE_CREATE_PROCESS, true);
			caps.setCapability(InternetExplorerDriver.IE_SWITCHES, "-private");
			new InternetExplorerDriver(caps);

			fileInputStream = new FileInputStream(props.getProperty("iexplorerprivate.excelPath"));
			browsername = "iexplorerprivate";
		} else if (browser.equals("safari")) {
			driver = new SafariDriver();
			fileInputStream = new FileInputStream(props.getProperty("safari.excelPath"));
			browsername = "safari";
		} else {

			System.setProperty("webdriver.chrome.driver", props.getProperty("chrome.webdriverPath"));
			//RemoteWebDriver driver1 = new RemoteWebDriver(new URL("http://<9.109.231.101>:<3166>"), DesiredCapabilities.chrome());

			ChromeOptions options = new ChromeOptions();
			DesiredCapabilities capabilities = DesiredCapabilities.chrome();
			capabilities.setCapability(ChromeOptions.CAPABILITY, options);
			driver = new ChromeDriver(capabilities);
			//driver = new ChromeDriver();
			fileInputStream = new FileInputStream(props.getProperty("chrome.excelPath"));
			browsername = "chrome";
		}

		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		// driver=new AppiumDriver();

	}

	@Test
	public void testAutorun() throws Exception {

		//Runtime.getRuntime().exec("C:\\SeleniumSetup\\Auto.exe");
		Properties props = new Properties();

		HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
		HSSFRow row;
		new LinkedHashMap<String, String>();
		workbook.getNumberOfSheets();
		int masterSheet = workbook.getSheetIndex("Master");
		workbook.getSheetName(masterSheet);
		// String SheetName=workbook.getSheetName(keywordsreference);
		int report = workbook.getSheetIndex("Reports");
		System.out.println("Report index "+report);
		workbook.getSheetName(report);
		HSSFSheet worksheet = workbook.getSheetAt(masterSheet);
		int rows;// no of rows
		rows = worksheet.getPhysicalNumberOfRows();
		File dir;
		HSSFSheet reportWorksheet = workbook.getSheetAt(3);
		// HSSFSheet keywordsreference=workbook.getSheetAt(4);
		HSSFSheet autoScript = workbook.getSheetAt(1);
		String screenshotFolder;

		String testCaseName = null;
		for (int i = 1; i < rows; i++) {

			row = worksheet.getRow(i);
			tobecontinue = true;
			System.out.println(row);
			if (row.getCell(3) != null) {
				String execStatus = row.getCell(3).getStringCellValue();
				if (null != execStatus && "yes".equalsIgnoreCase(execStatus)) {

					DateFormat dateFormat_1 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
					// get current date time with Calendar()
					Calendar cal = Calendar.getInstance();
					System.out.println(dateFormat_1.format(cal.getTime()));
					String Execution_Date_Time = dateFormat_1.format(cal.getTime());
					String[] TempString = Execution_Date_Time.split(" ");
					System.out.println(TempString[0]);
					System.out.println(TempString[1]);
					// get current start time with Date()
					DateFormat dateFormat2 = new SimpleDateFormat(" E yyyy/MM/dd 'at' HH:mm:ss");
					Date date2 = new Date();
					dateFormat2.format(date2);
					screenshotFolder = row.getCell(2).getStringCellValue();

					testCaseName = row.getCell(2).getStringCellValue();
					System.out.println(testCaseName);
					String TimeDate1 = TempString[0].replace("/", "-");
					String TimeDate2 = TempString[1].replace(":", ".");
					System.out.println(TimeDate1);
					System.out.println(TimeDate2);
					String UpdatedExecutionTimeDate = screenshotFolder + " executed " + "on" + " " + TimeDate1 + " "
							+ "at" + " " + TimeDate2;
					System.out.println("screenshots updating"+UpdatedExecutionTimeDate);
					System.out.println(UpdatedExecutionTimeDate);

					if (browsername.equalsIgnoreCase("firefox")) {
						new Properties();
						props.load(new FileInputStream(browserpath));
						dir = new File(props.getProperty("firefox.screenShotPath") + "\\" + UpdatedExecutionTimeDate);
						dir.mkdir();
					} else if (browsername.equalsIgnoreCase("iexplorer")) {
						// Properties props1 = new Properties();
						props.load(new FileInputStream(browserpath));
						dir = new File(props.getProperty("iexplorer.screenShotPath") + "\\" + UpdatedExecutionTimeDate);
						dir.mkdir();
					} else if (browsername.equalsIgnoreCase("safari")) {

						new Properties();
						props.load(new FileInputStream(browserpath));
						dir = new File(props.getProperty("safari.screenShotPath") + "\\" + UpdatedExecutionTimeDate);
						dir.mkdir();
					} else {
						new Properties();
						props.load(new FileInputStream(browserpath));
						dir = new File(props.getProperty("chrome.screenShotPath") + "\\" + UpdatedExecutionTimeDate);
						dir.mkdir();
					}

					testCaseName = row.getCell(2).getStringCellValue();
					int countOfTCs = i;

					DateFormat dateFormat1 = new SimpleDateFormat(" E yyyy/MM/dd 'at' HH:mm:ss");
					// get current date time with Date()
					Date date1 = new Date();
					System.out.println(dateFormat1.format(date1));
					// String dat;
					String start_dat = dateFormat1.format(date1);

					getTestExec(driver, workbook, testCaseName, countOfTCs, dir);

					DateFormat dateFormat = new SimpleDateFormat(" E yyyy/MM/dd 'at' HH:mm:ss");
					// get current date time with Date()
					Date date = new Date();
					System.out.println(dateFormat.format(date));
					// String dat;
					String dat = dateFormat.format(date);
					Date d1 = null;
					Date d2 = null;
					d1 = dateFormat.parse(start_dat);
					d2 = dateFormat.parse(dat);

					long diff = d2.getTime() - d1.getTime();
					System.out.println("totaltime=" + diff);
					long diffSeconds = diff / 1000 % 60;
					long diffMinutes = diff / (60 * 1000) % 60;
					long diffHours = diff / (60 * 60 * 1000) % 24;
					long diffDays = diff / (24 * 60 * 60 * 1000);
					long finaltimeinsec = diffSeconds + (diffMinutes * 60) + (diffHours * 60 * 60);

					System.out.print(String.valueOf(diffDays) + " days, ");
					System.out.print(diffHours + " hours, ");
					System.out.print(diffMinutes + " minutes, ");
					System.out.print(diffSeconds + " seconds.");
					// long exe_time=diffDays+diffHours+diffMinutes+diffSeconds;
					// System.out.println("totaltime="+exe_time);

					String time_min1 = "0" + String.valueOf(diffHours) + ":0" + String.valueOf(diffMinutes) + ":"
							+ String.valueOf(diffSeconds);

					generateExcelReport(workbook, autoScript, reportWorksheet, testCaseName, worksheet, time_min1, dat,
							start_dat, finaltimeinsec);

				}

			}

		}

	}

	/**
	 * get the cell value and process it case wise.
	 *
	 * @param driver
	 * @param workbook
	 * @param testCaseName
	 * @throws IOException
	 * @throws AWTException
	 * @throws HeadlessException
	 */
	// @Test
	// @Test(priority=2)
	private static void getTestExec(WebDriver driver, HSSFWorkbook workbook, String testCaseName, int countOfTCs,
			File dir) throws IOException, HeadlessException, AWTException {

		HSSFSheet worksheet = workbook.getSheetAt(1);
		System.out.println("The current worksheet is " + worksheet.getSheetName());
		HSSFSheet xpathWorksheet = workbook.getSheetAt(2);
		// HSSFWorkbook book=new HSSFWorkbook();
		// HSSFSheet sheet=book.createSheet("report");
		int rows = worksheet.getPhysicalNumberOfRows();
		int i = 1;
		// int r=0;
		while( (i < rows )&&(tobecontinue==true)) {

			HSSFRow row = worksheet.getRow(i);
			// System.out.println("row --" + row.getRowNum());
			if (row != null) {
				if (row.getCell(0) != null) {
					String testCaseStepName = row.getCell(0).getStringCellValue();
					if (null != testCaseStepName && testCaseStepName.equalsIgnoreCase(testCaseName)) {
						String actionDesc = null;
						String action = null;
						String xpath = null;
						String xpathId = null;
						String value = null;
						String value1 = null;
						if (row.getCell(2) != null)
							actionDesc = row.getCell(2).getStringCellValue();
						if (row.getCell(3) != null)
							action = row.getCell(3).getStringCellValue();
						if (row.getCell(4) != null) {
							xpath = row.getCell(4).getStringCellValue();
							if (!"".equals(xpath))
								xpathId = getXpathId(xpathWorksheet, xpath);
						}
						if (row.getCell(5) != null) {
							row.getCell(5).getCellType();
							HSSFCell cell = row.getCell(5);

							if (cell != null) {
								switch (cell.getCellType()) {
								case HSSFCell.CELL_TYPE_NUMERIC: {
									value = cell.getNumericCellValue() + "";
									break;
								}
								case HSSFCell.CELL_TYPE_STRING: {
									// STRING CELL TYPE
									value = cell.getRichStringCellValue().toString();

									break;
								}
								default: {
									// types other than String and Numeric.
									System.out.println("Type not supported.");
									break;
								}
								} // end switch
							} // value =
								// row.getCell(5).getRichStringCellValue().toString();

						}

						if (row.getCell(6) != null) {
							row.getCell(6).getCellType();
							HSSFCell cell = row.getCell(6);

							if (cell != null) {
								switch (cell.getCellType()) {
								case HSSFCell.CELL_TYPE_NUMERIC: {
									value1 = cell.getNumericCellValue() + "";
									break;
								}
								case HSSFCell.CELL_TYPE_STRING: {
									// STRING CELL TYPE
									value1 = cell.getRichStringCellValue().toString();

									break;
								}
								default: {
									// types other than String and Numeric.
									System.out.println("Type not supported.");
									break;
								}
								} // end switch
							} // value =
								// row.getCell(5).getRichStringCellValue().toString();

						}

						// getTestExec(workbook,testCaseName);
						System.out.println(
								"rows-- " + rows + " the actionDesc  " + actionDesc + " the action is " + action);
						if ("" != action && action.equalsIgnoreCase("LAUNCH_APPLICATION")) {
							HSSFSheet worksheetagain = workbook.getSheetAt(0);
							HSSFRow rowagain = worksheetagain.getRow(countOfTCs);
							String TC_URL = null;

							if ("" != (rowagain.getCell(4).getStringCellValue())) {
								TC_URL = rowagain.getCell(4).getStringCellValue();
								System.out.println(TC_URL);
								final long Stime = System.currentTimeMillis();
								launchApplication(workbook, row, driver, TC_URL, dir);
								final long Etime = System.currentTimeMillis();
								CalculateTime(workbook, row, driver, Stime, Etime);
							} else {

								String info = "Please enter the URL in master sheet!";
								JOptionPane.showMessageDialog(null, info);
								//driver.close();
								driver.quit();
							}
						}
						if ("" != action && action.equalsIgnoreCase("LAUNCH_URL")) {
							final long Stime = System.currentTimeMillis();

							launchURL(workbook, row, driver, value);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("VALIDATE_CLICK_WEBELEMENT")) {
							final long Stime = System.currentTimeMillis();
							validateClickElement(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action	&& (action.equalsIgnoreCase("CLICK_WEBELEMENT") || action.equals("CLICK_LINK"))) {
							final long Stime = System.currentTimeMillis();
							clickElement(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && (action.equalsIgnoreCase("CLICK_LINK_BY_TEXT"))) {
							final long Stime = System.currentTimeMillis();
							clickLink(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && (action.equalsIgnoreCase("DOUBLECLICK_WEBELEMENT"))) {
							final long Stime = System.currentTimeMillis();
							doubleClickElement(workbook, row, driver, xpathId);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SET_TEXTBOX")) {
							final long Stime = System.currentTimeMillis();
							setText(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SET_USERNAME")) {
							HSSFSheet worksheetagain = workbook.getSheetAt(0);
							HSSFRow rowagain = worksheetagain.getRow(countOfTCs);

							if ("" != (rowagain.getCell(5).getStringCellValue())) {
								String TC_Username = rowagain.getCell(5).getStringCellValue();
								final long Stime = System.currentTimeMillis();
								setUsername(workbook, row, driver, xpathId, TC_Username, dir);
								final long Etime = System.currentTimeMillis();
								CalculateTime(workbook, row, driver, Stime, Etime);

							} else {
								String info = "Please enter the username in master sheet!";
								JOptionPane.showMessageDialog(null, info);
								driver.quit();
							}
						} else if ("" != action && action.equalsIgnoreCase("SET_PASSWORD")) {
							HSSFSheet worksheetagain = workbook.getSheetAt(0);
							HSSFRow rowagain = worksheetagain.getRow(countOfTCs);
							if ("" != (rowagain.getCell(6).getStringCellValue())) {
								String TC_Password = rowagain.getCell(6).getStringCellValue();
								final long Stime = System.currentTimeMillis();
								setPassword(workbook, row, driver, xpathId, TC_Password, dir);
								final long Etime = System.currentTimeMillis();
								CalculateTime(workbook, row, driver, Stime, Etime);

							} else {
								String info = "Please enter the username in master sheet!";
								JOptionPane.showMessageDialog(null, info);
								driver.quit();
							}
						} else if ("" != action && action.equalsIgnoreCase("WAIT")) {
							final long Stime = System.currentTimeMillis();
							waitDriver(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("WAIT_SECONDS")) {
							final long Stime = System.currentTimeMillis();
							implicitWaitDriver(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("VALIDATE_EXIST")) {
							final long Stime = System.currentTimeMillis();
							validateExists(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("Timer")) {

							Timer(workbook, row, driver, xpathId, dir);
						} else if ("" != action && action.equalsIgnoreCase("CHECK_CHECKBOX")) {
							final long Stime = System.currentTimeMillis();
							checkCheckBox(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SCREENSHOT")) {
							final long Stime = System.currentTimeMillis();
							screenShot(workbook, driver, row, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SELECT_LISTBOX")) {
							final long Stime = System.currentTimeMillis();
							selectList(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("UNCHECK_CHECKBOX")) {
							final long Stime = System.currentTimeMillis();
							unCheckCheckBox(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SELECT_RADIOBUTTON_INDEXNUMBER")) {
							final long Stime = System.currentTimeMillis();
							selectRadioButton(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("CHECK_ENABLE_RADIOBUTTON_SELECT")) {
							final long Stime = System.currentTimeMillis();
							checkEnable(workbook, row, driver, xpathId, value, dir);

							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);

						} else if ("" != action && action.equalsIgnoreCase("CHECK_DISABLE")) {
							final long Stime = System.currentTimeMillis();
							checkDisable(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SCROLL_DOWN")) {
							final long Stime = System.currentTimeMillis();
							scrollDown(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("ZOOM")) {
							final long Stime = System.currentTimeMillis();
							zoom(workbook, row, driver, value);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("CLICKLINK_PROFILEBASED")) {
							final long Stime = System.currentTimeMillis();
							profileBased(workbook, row, driver, value);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("CHANGE_RESOLUTION")) {
							final long Stime = System.currentTimeMillis();
							changeResolution(workbook, row, driver, value);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("MAXIMIZE")) {
							final long Stime = System.currentTimeMillis();
							maximize(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("ALERT")) {
							final long Stime = System.currentTimeMillis();
							closeAlertAndGetItsText(workbook, row, driver, value);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("VALIDATE_TEXT_EXIST")) {
							final long Stime = System.currentTimeMillis();
							validateTextExists(workbook, row, driver, xpathId, value, value1, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("VALIDATE_TEXT_CONTAINS")) {
							final long Stime = System.currentTimeMillis();
							validateTextContains(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("Click_WebLink_By_Text")) {
							final long Stime = System.currentTimeMillis();
							ClickWebLinkByText(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("VALIDATE_LINK_BY_TEXT")) {
							final long Stime = System.currentTimeMillis();
							ValidateWebLinkByText(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("Scroll_Page")) {
							final long Stime = System.currentTimeMillis();
							ScrollPage(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("Select_Option_List")) {
							final long Stime = System.currentTimeMillis();
							SelectOptionList(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("Screen_Capture")) {
							final long Stime = System.currentTimeMillis();
							FullScreenCapture(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("CLICK_ELEMENT_BY_CSS")) {
							final long Stime = System.currentTimeMillis();
							ClickElementByCss(workbook, row, driver, xpathId);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("WAIT_TILL_VISIBLE")) {
							final long Stime = System.currentTimeMillis();
							WaitTillVisible(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("MOUSEHOVER_CLICK")) {
							final long Stime = System.currentTimeMillis();
							MouseHoverClick(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("FILE_UPLOAD")) {
							final long Stime = System.currentTimeMillis();
							fileupload(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SHOWPOPUPINFO")) {
							final long Stime = System.currentTimeMillis();
							ShowPopUpInfo(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SETTEXTBOXBYNAME")) {
							final long Stime = System.currentTimeMillis();
							setTextbyname(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SELECT_OPTION_BY_INDEX")) {
							final long Stime = System.currentTimeMillis();
							SelectOptionByIndex(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("GET_TEXT")) {
							final long Stime = System.currentTimeMillis();
							gettext(driver, workbook, row, xpathId);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("GET_DATE")) {
							final long Stime = System.currentTimeMillis();
							getdate(driver, workbook, row, xpathId);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("PAGE_REFRESH")) {
							final long Stime = System.currentTimeMillis();
							Pagerefresh(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("NEW_WINDOW")) {
							final long Stime = System.currentTimeMillis();
							windowhandles(workbook, row, driver, value, value1, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("NEW_WINDOW_PAGE")) {
							final long Stime = System.currentTimeMillis();
							windowhandles1(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("NEW_WINDOW_TITLE")) {
							final long Stime = System.currentTimeMillis();
							windowhandles1(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("NAVIGATE_PAGE")) {
							final long Stime = System.currentTimeMillis();
							navigate(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("FRAME")) {
							final long Stime = System.currentTimeMillis();
							frame(workbook, row, driver, value, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("TABLE_VALIDATE")) {
							final long Stime = System.currentTimeMillis();
							validatetable(workbook, row, driver, value, value1);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("CLOSE_CURRENT_URL")) {
							final long Stime = System.currentTimeMillis();
							closeurl(workbook, row, driver, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("VALIDATE_EXIST_CSS")) {
							final long Stime = System.currentTimeMillis();
							validateExistscss(workbook, row, driver, value, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("CLICK_TAB")) {
							final long Stime = System.currentTimeMillis();
							tab(workbook, row, driver, xpathId);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("SET_TEXT_OPTIONS")) {
							final long Stime = System.currentTimeMillis();
							selectOptionWithText(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("PACK_DISPLAYED")) {
							final long Stime = System.currentTimeMillis();
							pkgdisplayed(workbook, row, driver, xpathId, value, value1, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("CLEAR_HISTORY")) {
							final long Stime = System.currentTimeMillis();
							clearHistory(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("SET_PROXY")) {
							final long Stime = System.currentTimeMillis();
							setProxy(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("AutoIT_Download")) {
							final long Stime = System.currentTimeMillis();
							autoitdownloadpopup(workbook, row, driver, xpathId, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("DEFAULTFRAME")) {
							final long Stime = System.currentTimeMillis();
							defaultframe(workbook, row, driver, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("DisplayEnable")) {
							final long Stime = System.currentTimeMillis();
							displayEnable(workbook, row, driver, value, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("VALIDATE_NOT_EXIST")) {
							final long Stime = System.currentTimeMillis();
							validate_not_Exists(workbook, row, driver, value, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						} else if ("" != action && action.equalsIgnoreCase("GET_TEXT_DYNAMIC")) {
							final long Stime = System.currentTimeMillis();
							getTextDynamic(workbook, row, driver, xpathId, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);
						}

						else if ("" != action && action.equalsIgnoreCase("SET_TEXT_DYNAMIC")) {
							final long Stime = System.currentTimeMillis();
							setTextDynamic(workbook, row, driver,xpathId,dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);

						}

						else if ("" != action && action.equalsIgnoreCase("GET_DYNAMIC_URL_VALUE")) {
							final long Stime = System.currentTimeMillis();
							getDynamicUrlvalue(workbook, row, driver,value,dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);

						}
						else if ("" != action && action.equalsIgnoreCase("WINDOW_AUTHENTICATION")) {
							final long Stime = System.currentTimeMillis();
							windowAuthentication(workbook, row, driver, value, dir);
							final long Etime = System.currentTimeMillis();
							CalculateTime(workbook, row, driver, Stime, Etime);

						}



						else {
							System.out.println("Not a listed action");
						}

					}
				}

			}
			i++;

		}

	}



	private static String getXpathId(HSSFSheet xpathWorksheet, String xpathId) {

		String xPathId = null;
		int i = 1;
		int rows = xpathWorksheet.getPhysicalNumberOfRows();
		while (i < rows) {
			// System.out.println("i --" + i);
			HSSFRow row = xpathWorksheet.getRow(i);

			if (row != null) {
				if (row.getCell(1) != null) {
					if (xpathId.equalsIgnoreCase(row.getCell(1).getStringCellValue())) {
						xPathId = row.getCell(2).getStringCellValue();
					}
				}
			}
			i++;
		}
		return xPathId;
	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void Pagerefresh(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {
			driver.navigate().refresh();


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}

	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void SelectOptionByIndex(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String Id,
			String value, File dir) throws IOException {
		try {
			int index = Integer.parseInt(value);
			WebElement ltbox = driver.findElement(By.id(Id));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					ltbox);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", ltbox);

			Select sec = new Select(ltbox);
			sec.selectByIndex(index);


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			tobecontinue = false;
		}

	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void setTextbyname(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			WebElement userId = driver.findElement(By.name(xpathId));

			userId.clear();
			Actions builder = new Actions(driver);
			Actions seriesOfActions = builder.moveToElement(userId).click().sendKeys(userId, value);
			seriesOfActions.perform();

			// userId.sendKeys(value);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	// performing tab operation in keyboard

	private static void tab(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId) throws IOException {
		try {

			WebElement e = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", e);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", e);

			Actions action = new Actions(driver);
			action.sendKeys(e, Keys.TAB).build().perform();

			System.out.println("e");
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
		}
	}

	/**n
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void setPassword(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {
			WebElement userId = driver.findElement(By.xpath(xpathId));

			userId.clear();
			Actions builder = new Actions(driver);
			Actions seriesOfActions = builder.moveToElement(userId).click().sendKeys(userId, value);
			seriesOfActions.perform();

			// userId.sendKeys(value);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void setUsername(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {
			WebElement userId = driver.findElement(By.xpath(xpathId));

			userId.clear();
			Actions builder = new Actions(driver);
			Actions seriesOfActions = builder.moveToElement(userId).click().sendKeys(userId, value);
			seriesOfActions.perform();

			// userId.sendKeys(value);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void setText(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {

			WebElement userId = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					userId);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", userId);

			// clearing the field
			userId.clear();
			// sending the value to the textfield
			userId.sendKeys(value);

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	public static void selectOptionWithText(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		new WebDriverWait(driver, 5);
		try {
			WebElement autoOptions = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					autoOptions);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", autoOptions);

			autoOptions.clear();
			autoOptions.sendKeys(value);
			autoOptions.sendKeys(Keys.BACK_SPACE);
			autoOptions.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(2000);
			autoOptions.click();


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * set the test box value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void ShowPopUpInfo(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {
			String info = value;
			JOptionPane.showMessageDialog(null, info);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}

	}

	 /**
	 * Mouse Hover AND then click
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void MouseHoverClick(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			WebElement mnuElement;
			// WebElement mnuElement1;

			mnuElement = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					mnuElement);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", mnuElement);

			Actions builder = new Actions(driver);
			// Move cursor to the Main Menu Element

			builder.moveToElement(mnuElement).perform();
			// Giving 5 Secs for submenu to be displayed
			// Thread.sleep(1000L);
			// Clicking on the Hidden SubMenu
			// driver.findElement(By.linktext(value)).click;

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	private static void setProxy(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {

			String usedProxy = "http://fastweb.int.bell.ca:8080";

			org.openqa.selenium.Proxy proxy = new org.openqa.selenium.Proxy();
			proxy.setHttpProxy(usedProxy).setFtpProxy(usedProxy).setSslProxy(usedProxy);
			DesiredCapabilities cap = new DesiredCapabilities();
			cap.setCapability(CapabilityType.PROXY, proxy);

			driver = new FirefoxDriver(cap);

			// WebElement mnuElement1;
			/*
			 * driver=null; org.openqa.selenium.Proxy proxy = new
			 * org.openqa.selenium.Proxy();
			 * proxy.setSslProxy("fastweb.int.bell.ca"+":"+8083);
			 * proxy.setFtpProxy("fastweb.int.bell.ca"+":"+8083);
			 *
			 *
			 * DesiredCapabilities dc = DesiredCapabilities.firefox();
			 * dc.setCapability(CapabilityType.PROXY, proxy);
			 *
			 */
			/*
			 * FirefoxProfile profile = new FirefoxProfile();
			 * profile.setPreference("network.proxy.type", 1);
			 * profile.setPreference("network.proxy.http",
			 * "fastweb.int.bell.ca");
			 * profile.setPreference("network.proxy.http_port", 8083);
			 * profile.setPreference("network.proxy.ssl",
			 * "fastweb.int.bell.ca");
			 * profile.setPreference("network.proxy.ssl_port", 8083); driver=new
			 * FirefoxDriver(profile);
			 */
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 *
	 * click the element ex: button.
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void clickElement(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {

			WebElement query1 = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					query1);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", query1);
			js.executeScript("arguments[0].click();", query1);


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
			//Autorun auto = new Autorun();
			/*driver.quit();
			haltScript();*/
		}

	}

	private static void autoitdownloadpopup(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {

			Runtime.getRuntime().exec(value);

			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}

	}

	/**
	 * click link
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void clickLink(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {

			WebElement query = driver.findElement(By.partialLinkText(value));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					query);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", query);

			query.click();
			//js.executeScript("arguments[0].setAttribute('style','background: green; border: 4px solid green;'); ",query);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * at times in IE some elements needs double click hence implemeneted.
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void doubleClickElement(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId)
			throws IOException {
		try {
			WebElement query = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					query);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", query);

			query.click();
			query.click();

					row.createCell(7).setCellValue("Pass");
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * launch the url in specific browser
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void launchApplication(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {


			driver.get(value);


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");

			screenShot_failure(workbook, driver, row, value, dir);
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	/**
	 * launch the url in specific browser
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void launchURL(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value)
			throws IOException {
		try {

			driver.get(value);

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");

			setCell(workbook, row, "Pass");
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

	// Getting dynamic Value from webpage

	private static String getTextDynamic(HSSFWorkbook workbook, HSSFRow row, WebDriver driver,
			String xpathId, File dir) throws IOException {
		try {
			System.out.println("dynamic value is going to display");

		driver.findElement(By.xpath(xpathId)).click();

		System.out.println("after click");

		WebElement Ele = driver.findElement(By.xpath(xpathId));
			String Temp= Ele.getText().toString();

		System.out.println(Temp);

			TxtID = Temp.substring(17).trim();

			System.out.println(TxtID);

			setCell(workbook, row, "Pass");
			return TxtID;

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			printErrorMessage(row, e);
			tobecontinue = false;
		}
		return TxtID;

	}
	// writing Dynamic text value fetched

	public static void setTextDynamic(HSSFWorkbook workbook, HSSFRow row,
			WebDriver driver,String xpathId, File dir) throws IOException {
		try {

			WebElement Ele = driver.findElement(By.xpath(xpathId));

			Ele.clear();

			Ele.sendKeys(TxtID);
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			printErrorMessage(row, e);
			tobecontinue = false;
		}

	}

// getting the dynamic id in the url using particular value to trim the string

	public static void getDynamicUrlvalue(HSSFWorkbook workbook, HSSFRow row,
			WebDriver driver, String value, File dir) throws IOException {
		try {

			String Curl = driver.getCurrentUrl();

			System.out.println(Curl);

			StringTokenizer stok = new StringTokenizer(Curl, "=");
			String str = new String();
			while (stok.hasMoreTokens()) {
				str = (String) stok.nextElement();


			}
			System.out.println(str);
			TxtID = str;

			System.out.println(TxtID);


			StringTokenizer stok1= new StringTokenizer(value, "=");

			String str1=stok1.nextToken();

			System.out.println(str1);
			String str2=stok1.nextToken();
			System.out.println(str2);
			String str3=stok1.nextToken();

			System.out.println(str3);



			String Newurl= str1+"="+TxtID+"&BTCID"+"="+str3;

			System.out.println(Newurl);

			driver.get(Newurl);


			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			printErrorMessage(row, e);
		}

	}



	/**
	 * check the checkbox
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	public static void checkCheckBox(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {
			boolean checkStatus = driver.findElement(By.xpath(xpathId)).isSelected();
			// boolean checkStatus =
			// driver.findElement(By.id(xpathId)).isSelected();
			WebElement ele = driver.findElement(By.xpath(xpathId));
			if (checkStatus == false)

				((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", ele);
			ele.click();

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
		printErrorMessage(row, e);
		tobecontinue = false;
		}
	}

	public static void displayEnable(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {

			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("document.getElementByXpath(xpathId).style.display='block';");

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			printErrorMessage(row, e);
		}
	}

	/**
	 * screenshot
	 *
	 * @param driver
	 * @param value
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws InterruptedException
	 */
	public static void screenShot(HSSFWorkbook workbook, WebDriver driver, HSSFRow row, String value, File dir)
			throws FileNotFoundException, IOException {

		File ScreenshotsFolder = dir;
		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		// get current date time with Calendar()
		Calendar cal = Calendar.getInstance();
		System.out.println(dateFormat.format(cal.getTime()));
		String Execution_Date_Time = dateFormat.format(cal.getTime());
		String[] TempString = Execution_Date_Time.split(" ");
		System.out.println(TempString[0]);
		System.out.println(TempString[1]);
		String TimeDate1 = TempString[0].replace("/", "-");
		String TimeDate2 = TempString[1].replace(":", ".");
		System.out.println(TimeDate1);
		System.out.println(TimeDate2);
		//temp.clear();
		String UpdatedExecutionTimeDate = value + " taken " + "on" + " " + TimeDate1 + " " + "at" + " " + TimeDate2;
		System.out.println(UpdatedExecutionTimeDate);
		String tcName = value;
		Properties props = new Properties();
		props.load(new FileInputStream(browserpath));
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		if (browsername.equalsIgnoreCase("firefox")) {
			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile,new
			// File(ScreenshotsFolder+"//"+tcName+"//"+UpdatedExecutionTimeDate+".png"));

		} else if (browsername.equalsIgnoreCase("iexplorer")) {
			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("iexplorer.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		} else if (browsername.equalsIgnoreCase("safari")) {
			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("safari.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		} else {

			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("chrome.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		}
		try {

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			System.out.println("Fail to update the screenshot");
			tobecontinue = false;

			// printErrorMessage(row, e);
		}
	}

	public static void screenShot_failure(HSSFWorkbook workbook, WebDriver driver, HSSFRow row, String value, File dir)
			throws FileNotFoundException, IOException {

		File ScreenshotsFolder = dir;
		String tcName = value;
		String failedscreen = "failstepscreen";
		Properties props = new Properties();
		props.load(new FileInputStream(browserpath));
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		if (browsername.equalsIgnoreCase("firefox")) {

			FileUtils.copyFile(scrFile,
					new File(ScreenshotsFolder + "//" + failedscreen + "_" + row.getRowNum() + ".png"));

		} else if (browsername.equalsIgnoreCase("iexplorer")) {
			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("iexplorer.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		} else if (browsername.equalsIgnoreCase("safari")) {
			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("safari.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		} else {

			FileUtils.copyFile(scrFile, new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
			// FileUtils.copyFile(scrFile, new
			// File(props.getProperty("chrome.screenShotPath")+"\\"
			// +tcName+"_"+row.getRowNum()+".png"));
		}
			}

	public static void FullScreenCapture(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws AWTException, IOException {
		File ScreenshotsFolder = dir;
		String tcName = value;
		Properties props = new Properties();
		props.load(new FileInputStream(browserpath));
		BufferedImage image = new Robot()
				.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		if (browsername.equalsIgnoreCase("firefox")) {

			ImageIO.write(image, "png", new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
		} else if (browsername.equalsIgnoreCase("iexplorer")) {

			ImageIO.write(image, "png", new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
		} else if (browsername.equalsIgnoreCase("safari")) {

			ImageIO.write(image, "png", new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
		} else if (browsername.equalsIgnoreCase("chrome")) {

			ImageIO.write(image, "png", new File(ScreenshotsFolder + "//" + tcName + "_" + row.getRowNum() + ".png"));
		} else if(browsername.equalsIgnoreCase("navigator")){

			ImageIO.write(image,  "png", new File(ScreenshotsFolder + "//" + tcName + "_" +row.getRowNum() + ".png"));
		}
		try {

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			System.out.println("Failed to udpate the case");
			// printErrorMessage(row, e);

		}
	}

	/**
	 * uncheck the checkbox
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	public static void unCheckCheckBox(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			System.out.println("xpathid --- " + xpathId);
			WebElement checkStatus = driver.findElement(By.xpath(xpathId));
			System.out.println("checkStatus --- " + checkStatus);
			// if (checkStatus!=null)
			driver.findElement(By.xpath(xpathId)).click();
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);


			tobecontinue = false;
		}
	}

	/**
	 * uncheck the checkbox
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	public static void WaitTillVisible(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			System.out.println("xpathid --- " + xpathId);
			WebDriverWait wait = new WebDriverWait(driver, 40);

			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpathId)));
			//driver.manage().timeouts().implicitlyWait(150, TimeUnit.SECONDS);

			// driver.findElement(By.xpath(xpathId)).isDisplayed();

			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");

			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * select the radio button
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void selectRadioButton(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			// int index = Integer.parseInt(value);

			WebElement ele = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", ele);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", ele);

			ele.click();

			// radioGroup.get(index).click();
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * wait in the secs
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void implicitWaitDriver(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {
			// driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			// int time=Integer.parseInt(value);
			// driver.manage().timeouts().implicitlyWait(20000,TimeUnit.SECONDS);
			String TIME = value;
			// driver.manage().timeouts().implicitlyWait(Integer.parseInt(TIME),TimeUnit.SECONDS);
			Long waitcheck = Long.parseLong(TIME);
			Thread.sleep(waitcheck);

			setCell(workbook, row, "Pass");
			// System.out.println("i ma waiting..");
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * profile based click based on text like TV, Mobile
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void profileBased(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value)
			throws IOException {
		try {
			java.util.List<WebElement> aElements = driver.findElements(By.tagName("a"));
			for (WebElement element : aElements) {
				if (element.getText().startsWith(value))
					element.click();
			}
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	private static void windowhandles1(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {

		try {

			Set<String> allWindows = driver.getWindowHandles();
			/*
			 * Iterator i= allWindows.iterator();
			 *
			 Set<String> allWindows =driver.getWindowHandles();
			 *
			 * String ParentWindow = (String) i.next(); String ChildWindow =
			 * (String) i.next(); System.out.println(ParentWindow);
			 * System.out.println(ChildWindow);
			 * driver.switchTo().window(ChildWindow); setCell(workbook, row,
			 * "Pass");
			 */

			for (String currentWindow : allWindows) {
				driver.switchTo().window(currentWindow);

				System.out.println(currentWindow);
				driver.manage().window().maximize();
				String url = driver.getCurrentUrl();

				String title = driver.getTitle();

				System.out.println(title);

				System.out.println(url);


				setCell(workbook, row, "Pass");
			}

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
			System.out.println("navigated to the new window");
		}

	}

	// --------------window handles--------------

	private static void windowhandles(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, String value1,
			File dir) throws IOException {
		try {

			Set<String> allWindows = driver.getWindowHandles();
			Iterator<String> it = allWindows.iterator();
			//iterate through your windows
			while (it.hasNext()){
			String parent = it.next();
			System.out.println(parent);
			String newwin = it.next();
			System.out.println(newwin);
				System.out.println(allWindows);
			}
			for (String currentWindow : allWindows) {
				driver.switchTo().window(currentWindow);
				System.out.println(currentWindow);
			}

			Thread.sleep(30000);

			String url = driver.getCurrentUrl();
			System.out.println(url);
			// String [] strArray=url.split("//mybell.bell");
			// String [] strArray =url.replaceFirst(value, value1);
			String newurl = url.replaceFirst(value, value1);
			// String newurl=strArray[0]+value+strArray[1];
			System.out.println(newurl);
			// navigate(workbook,row,driver,value,dir);
			launchURL(workbook, row, driver, newurl);

			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	private static void navigate(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {

			driver.navigate().to(value);

			// driver.manage().window().maximize();
			driver.manage().window().maximize();

			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	private static void closeurl(HSSFWorkbook workbook, HSSFRow row, WebDriver driver,String value, File dir)
			throws IOException {
		try {

			Thread.sleep(5000);
			/*
			 * DesiredCapabilities caps = new DesiredCapabilities();
			 * caps.setCapability(CapabilityType.ForSeleniumServer.
			 * ENSURING_CLEAN_SESSION,true);
			 */
			driver.close();
			driver.manage().deleteAllCookies();


			Thread.sleep(5000);


			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}

	private static void frame(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, String xpathId,
			File dir) throws IOException {
		try {

			driver.switchTo().frame(value);


			// driver.findElement(By.xpath(xpathId)).click();
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	private static void defaultframe(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, File dir)
			throws IOException {
		try {

			driver.switchTo().defaultContent();


			// driver.findElement(By.xpath(xpathId)).click();
			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * wait till the element is visible
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void waitDriver(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, final String xpathId,
			String value, File dir) throws IOException {
		try {
			(new WebDriverWait(driver, 1000)).until(new ExpectedCondition<WebElement>() {
				@Override
				public WebElement apply(WebDriver d) {
					return d.findElement(By.xpath(xpathId));
				}
			});


			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * validate if element exists
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void validateExists(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			WebElement ele = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", ele);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", ele);
			boolean isPresent = ele.isDisplayed();

			if (isPresent)
				setCell(workbook, row, "Pass");
			else
				setCell(workbook, row, "Fail");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	private static void validate_not_Exists(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			boolean isPresent = driver.findElement(By.xpath(xpathId)).isDisplayed();

			// Heighlate the element
			/*
			 * JavascriptExecutor js=(JavascriptExecutor)driver;
			 * js.executeScript(
			 * "arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');"
			 * , ele); Thread.sleep(200); js.executeScript(
			 * "arguments[0].setAttribute('style','border: solid 2px white')",
			 * ele);
			 */

			if (isPresent)
				setCell(workbook, row, "Fail");
			else
				setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * Note the time taken by system to navigate from one step to another
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void Timer(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, File dir)
			throws IOException {
		try {
			boolean isPresent = false;
			final long startTime = System.currentTimeMillis();
			System.out.println(startTime);
			while (!isPresent) {
				System.out.println("m here in while loop");
				isPresent = driver.findElement(By.xpath(xpathId)).isDisplayed();
			}
			final long endTime = System.currentTimeMillis();

			System.out.println(endTime);
			System.out.println("Total execution time: " + (endTime - startTime));

			System.out.println(String.format("%d hr,%d min, %d sec, %d milli sec",
					TimeUnit.MILLISECONDS.toHours(endTime - startTime),
					TimeUnit.MILLISECONDS.toMinutes(endTime - startTime)
							- TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(endTime - startTime)), // The
																											// change
																											// is
																											// in
																											// this
																											// line
					TimeUnit.MILLISECONDS.toSeconds(endTime - startTime)
							- TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(endTime - startTime)),
					(endTime - startTime)
							- TimeUnit.SECONDS.toMillis(TimeUnit.MILLISECONDS.toSeconds(endTime - startTime))));

			if (isPresent) {

				setCell(workbook, row, "Pass");
				try {
					Properties props = new Properties();
					props.load(new FileInputStream(browserpath));
					// Thread.sleep(3000);
					row.createCell(7)
							.setCellValue(String.format("%d hr,%d min, %d sec, %d milli sec",
									TimeUnit.MILLISECONDS.toHours(endTime - startTime),
									TimeUnit.MILLISECONDS.toMinutes(endTime - startTime) - TimeUnit.HOURS
											.toMinutes(TimeUnit.MILLISECONDS.toHours(endTime - startTime)), // The
																											// change
																											// is
																											// in
																											// this
																											// line
							TimeUnit.MILLISECONDS.toSeconds(endTime - startTime)
									- TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(endTime - startTime)),
							(endTime - startTime)
									- TimeUnit.SECONDS.toMillis(TimeUnit.MILLISECONDS.toSeconds(endTime - startTime))));
					fileInputStream.close();
					//
					if (browsername.equals("firefox")) {
						// Thread.sleep(3000);
						outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));
					} else if (browsername.equals("iexplorer")) {

						outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
					} else if (browsername.equals("chrome")) {

						outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
					} else {
						outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
					}

					workbook.write(outfile);

					outfile.close();
				} catch (Exception e) {

					printErrorMessage(row, e);
					e.printStackTrace();

				}

			} else {
				setCell(workbook, row, "Fail");
				HSSFCell cell = row.createCell(8);
				cell.setCellValue("Element is not present");
			}
		} catch (Exception e) {
			setCell(workbook, row, "Fail");

			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	private static void validateExistscss(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {
			boolean isPresent = driver.findElement(By.cssSelector(xpathId)).isDisplayed();
			if (isPresent)
				setCell(workbook, row, "Pass");
			else
				setCell(workbook, row, "Fail");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * validates if element exists with the same text value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void validateTextExists(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, String value1, File dir) throws IOException {
		try {
			WebElement ele = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", ele);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", ele);

			String text = ele.getText();
			System.out.println(text);

			if (text.contentEquals(value))

				setCell(workbook, row, "Pass");
			else {
				setCell(workbook, row, "Fail");
				HSSFCell cell = row.createCell(8);
				cell.setCellValue("text doesnot matches");
			}
			// screenShot_failure(workbook,driver, row, value,dir);
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}

	private static void printErrorMessage(HSSFRow row, Exception e) {
		HSSFCell cell = row.createCell(8);
		cell.setCellValue(e.getMessage());

	}

	/**
	 * validates if element exists with the same text value
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void validateTextContains(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) throws IOException {
		try {

			WebElement ele = driver.findElement(By.xpath(xpathId));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');", ele);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", ele);

			String text = ele.getText();
			System.out.println(text);
			// if (text.equalsIgnoreCase(value))
			if (text.contains(value))

				setCell(workbook, row, "Pass");
			else
				setCell(workbook, row, "Fail");
			HSSFCell cell = row.createCell(8);
			cell.setCellValue("Text does not exist");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	@AfterMethod
	public static void tearDown() throws Exception {
		//driver.quit();

		String verificationErrorString = verificationErrors.toString();
		if (!"".equals(verificationErrorString)) {
			fail(verificationErrorString);
		}
	}

	public static void closeAlertAndGetItsText(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value)
			throws IOException {
		// try {
		Alert alert = driver.switchTo().alert();

		if (value == "YES") {
			alert.accept();
			setCell(workbook, row, "Pass");
		} else if (value == "NO") {

			alert.accept();
			setCell(workbook, row, "Pass");

		}

		// String alertText = alert.getText();
		// if (acceptNextAlert) {
		// alert.accept();
		// } else {
		// alert.dismiss();
		// }
		// return alertText;
		// } finally {
		// acceptNextAlert = true;
		// }
	}

	/**
	 * set textbox
	 *
	 * @param workbook
	 * @param row
	 * @param val
	 * @throws IOException
	 */
	// @Test
	// @Parameters("browsername")
	// @Test(priority = 3, dependsOnMethods = "getTestExec")
	// @AfterMethod
	private static void setCell(HSSFWorkbook workbook, HSSFRow row, String val) throws IOException {
		try {
			Properties props = new Properties();
			props.load(new FileInputStream(browserpath));
			// Thread.sleep(3000);
			row.createCell(6).setCellValue(val);
			fileInputStream.close();
			// if(browsername.equals("firefox") || browsername.equals("chrome")
			// || browsername.equals("iexplorer"))
			// {
			// outfile = new FileOutputStream(
			// props.getProperty("firefox.excelPath"));
			// outfile = new FileOutputStream(
			// props.getProperty("iexplorer.excelPath"));
			// outfile = new FileOutputStream(
			// props.getProperty("chrome.excelPath"));
			// }
			if (browsername.equals("firefox")) {
				// Thread.sleep(3000);
				outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));
			} else if (browsername.equals("iexplorer")) {

				outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
			} else if (browsername.equals("chrome")) {

				outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
			} else {
				outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
			}

			workbook.write(outfile);

			outfile.close();
		} catch (Exception e) {
			if ("Fail".equals(val))
				printErrorMessage(row, e);
			e.printStackTrace();

		}

	}



	// ------------gettextandwritinginexcel---------

	private static void gettext(WebDriver driver, HSSFWorkbook workbook, HSSFRow row, String xpathId)
			throws IOException {
		try {
			String val = driver.findElement(By.xpath(xpathId)).getText();

			// DateFormat dateFormat1 = new SimpleDateFormat(" E yyyy/MM/dd 'at'
			// HH:mm:ss");
			// Date date1 = new Date();
			// if(val.equals(date1))
			// {
			// String dat =dateFormat1.format(date1);
			//
			//
			// }
			// else
			// {
			// val.split("manual");
			Properties props = new Properties();
			props.load(new FileInputStream(browserpath));
			row.createCell(9).setCellValue(val);

			// String val1=
			// driver.findElement(By.xpath(".//*[@id='printContentArea']/div/table[2]/tbody/tr[2]/td[2]")).getText();
			// DateFormat dateFormat = new SimpleDateFormat(" MM/dd/yyyy");
			// get current date time with Date()
			// Date date = new Date(val);
			// System.out.println(dateFormat.format(date));

			// String dat;
			// String dat =dateFormat.format(date);
			// Date d1 = null;
			// Date d2 = null;
			// d1 = dateFormat.parse(dat);
			// row.createCell(10).setCellValue(dat);

			// String val2=
			// driver.findElement(By.xpath(".//*[@id='printContentArea']/div/table[2]/tbody/tr[3]/td[2]")).getText();

			// row.createCell(11).setCellValue(val2);
			// if(dat.compareTo("Sat")==0|| dat.compareTo("Sun")==0)
			// {
			// row.createCell(11).setCellValue("weekend");
			// }
			// else
			// {
			// row.createCell(11).setCellValue("weekday");
			// }

			// CellStyle cellStyle = workbook.createCellStyle();
			// CreationHelper createHelper = workbook.getCreationHelper();
			// cellStyle.setDataFormat(
			// createHelper.createDataFormat().getFormat("E MM d, yyyy h:mm"));
			// // BuiltinFormats.getBuiltinFormat(d1);
			// Cell cell=row.createCell(10);
			// cell = row.createCell(10);
			// cell.setCellValue(d1);
			// cell.setCellStyle(cellStyle);
			// row.createCell(10).setCellValue(d1);

			// System.out.println("converted date"+d1);

			// d2 = dateFormat.parse(dat);

			// DateFormat formatter = null;
			// Date convertedDate = null;
			// String dMMMMyy = val;
			// formatter = new SimpleDateFormat("MMMM,DDyy");
			// convertedDate = (Date) formatter.parse(dMMMMyy);
			// System.out.println("Date from dd-MMMM-yy String in Java : " +
			// convertedDate);
			//

			fileInputStream.close();
			if (browsername.equals("firefox")) {
				outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));
			} else if (browsername.equals("iexplorer")) {
				outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
			} else if (browsername.equals("safari")) {
				outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
			} else {
				outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
			}
			workbook.write(outfile);
			outfile.close();
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			// if("Fail".equals(val))
			setCell(workbook, row, "fail");
			printErrorMessage(row, e);
			e.printStackTrace();

		}

	}

	@SuppressWarnings("deprecation")
	private static void getdate(WebDriver driver, HSSFWorkbook workbook, HSSFRow row, String xpathId)
			throws IOException {
		try {
			// String val= driver.findElement(By.xpath(xpathId)).getText();

			// DateFormat dateFormat1 = new SimpleDateFormat(" E yyyy/MM/dd 'at'
			// HH:mm:ss");
			// Date date1 = new Date();
			// if(val.equals(date1))
			// {
			// String dat =dateFormat1.format(date1);
			//
			//
			// }
			// else
			// {
			// val.split("manual");

			// row.createCell(9).setCellValue(val);

			String val1 = driver.findElement(By.xpath(xpathId)).getText();
			new SimpleDateFormat("  MM/dd/yyyy");
			new Date(val1);

			// String dat;
			// String dat =dateFormat.format(date);
			// Date d1 = null;
			// Date d2 = null;
			// d1 = dateFormat.parse(dat);
			row.createCell(9).setCellValue(val1);

			Properties props = new Properties();
			props.load(new FileInputStream(browserpath));

			// String val2=
			// driver.findElement(By.xpath(".//*[@id='printContentArea']/div/table[2]/tbody/tr[3]/td[2]")).getText();

			// row.createCell(11).setCellValue(val2);
			// if(dat.compareTo("Sat")==0|| dat.compareTo("Sun")==0)
			// {
			// row.createCell(11).setCellValue("weekend");
			// }
			// else
			// {
			// row.createCell(11).setCellValue("weekday");
			// }

			// CellStyle cellStyle = workbook.createCellStyle();
			// CreationHelper createHelper = workbook.getCreationHelper();
			// cellStyle.setDataFormat(
			// createHelper.createDataFormat().getFormat("E MM d, yyyy h:mm"));
			// // BuiltinFormats.getBuiltinFormat(d1);
			// Cell cell=row.createCell(10);
			// cell = row.createCell(10);
			// cell.setCellValue(d1);
			// cell.setCellStyle(cellStyle);
			// row.createCell(10).setCellValue(d1);

			// System.out.println("converted date"+d1);

			// d2 = dateFormat.parse(dat);

			// DateFormat formatter = null;
			// Date convertedDate = null;
			// String dMMMMyy = val;
			// formatter = new SimpleDateFormat("MMMM,DDyy");
			// convertedDate = (Date) formatter.parse(dMMMMyy);
			// System.out.println("Date from dd-MMMM-yy String in Java : " +
			// convertedDate);
			//

			fileInputStream.close();
			if (browsername.equals("firefox")) {
				outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));
			} else if (browsername.equals("iexplorer")) {
				outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
			} else if (browsername.equals("safari")) {
				outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
			} else {
				outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
			}
			workbook.write(outfile);
			outfile.close();
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			// if("Fail".equals(val))
			setCell(workbook, row, "fail");
			printErrorMessage(row, e);
			e.printStackTrace();

		}

	}

	/**
	 * validate if the element exist and then click
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 */
	private static void ClickElementByCss(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String Css) {
		try {
			boolean isPresent = driver.findElement(By.cssSelector(Css)).isDisplayed();
			WebElement query = driver.findElement(By.cssSelector(Css));
			if (isPresent)
				query.click();
		} catch (Exception e) {

			e.printStackTrace();
		}
	}

	/**
	 * validate if the element exist and then click
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 */
	private static void validateClickElement(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId,
			String value, File dir) {
		try {
			boolean isPresent = driver.findElement(By.xpath(xpathId)).isDisplayed();
			if (isPresent)
				clickElement(workbook, row, driver, xpathId, value, dir);
		} catch (Exception e) {

			e.printStackTrace();
			tobecontinue = false;
		}
	}

	/**
	 * click an element after finding it by text
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param TextLink
	 * @throws IOException
	 */
	private static void ValidateWebLinkByText(HSSFWorkbook workbook, HSSFRow row, WebDriver driver,
			final String TextLink, String value, File dir) throws IOException {
		try {
			boolean isPresent = driver.findElement(By.linkText(TextLink)).isDisplayed();

			if (isPresent)
				setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * click an element after finding it by text
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param TextLink
	 * @throws IOException
	 */
	private static void ClickWebLinkByText(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, final String TextLink,
			String value, File dir) throws IOException {
		try {
			WebElement isPresent = (new WebDriverWait(driver, 900)).until(new ExpectedCondition<WebElement>() {
				@Override
				public WebElement apply(WebDriver d) {
					return d.findElement(By.linkText(TextLink));
				}
			});
			isPresent.click();
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	private static void clearHistory(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {

			Process p = Runtime.getRuntime().exec("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255");
			p.waitFor();
			Thread.sleep(5000);
			/*
			 * DesiredCapabilities caps = new DesiredCapabilities();
			 * caps.setCapability(CapabilityType.ForSeleniumServer.
			 * ENSURING_CLEAN_SESSION,true);
			 */
			driver.manage().deleteAllCookies();


			Thread.sleep(5000);


			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}

	/**
	 * Scroll the page till the specified element
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @throws IOException
	 */
	private static void ScrollPage(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {
			WebElement element = driver.findElement(By.xpath(xpathId));

			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
			Thread.sleep(500);

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					element);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", element);

			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * Select the option by visible text
	 *
	 * @param workbook
	 * @param row
	 * @ param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void SelectOptionList(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String elementname,
			String value, File dir) throws IOException {
		try {
			Select select = new Select(driver.findElement(By.name(elementname)));

			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					select);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", select);

			select.selectByVisibleText(value);


			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * select list
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void selectList(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {
			// int val=Integer.parseInt(value);
			Select selectBox = new Select(driver.findElement(By.xpath(xpathId)));
			WebElement select = driver.findElement(By.xpath(xpathId));
			// Heighlate the element
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
					select);
			Thread.sleep(200);
			js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", select);

			// selectBox.selectByIndex(val);
			selectBox.selectByValue(value);

			setCell(workbook, row, "Pass");
		}

		catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			printErrorMessage(row, e);
			tobecontinue = false;

		}

	}

	/**
	 * check if element diabled.
	 *
	 * @param workbook
	 * @param row
	 * @param driverv
	 * @param xpathId
	 */

	/*
	 * public static void productDisplayed(String productid, WebDriver driver,
	 * boolean isdiplay) throws Exception { By product=
	 * By.xpath("//a[@id='"+productid+"']"); try{ isdiplay=
	 * driver.findElement(product).isDisplayed();
	 *
	 * } catch (Exception e){
	 *
	 * } }
	 */

	private static void checkDisable(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) {
		try {
			boolean isDisabled = driver.findElement(By.xpath(xpathId)).isEnabled();
			if (isDisabled == true) {

				setCell(workbook, row, "Pass");

			} else {

				setCell(workbook, row, "fail");
				screenShot_failure(workbook, driver, row, value, dir);

			}
		} catch (IOException e) {
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * check if element is enabled.
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 */
	private static void checkEnable(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) {
		try {

			boolean isEnabled = driver.findElement(By.xpath(xpathId)).isEnabled();

			if (isEnabled == true) {
				WebElement query = driver.findElement(By.xpath(xpathId));

				// Heighlate the element
				JavascriptExecutor js = (JavascriptExecutor) driver;
				js.executeScript("arguments[0].setAttribute('style', 'background: yellow; border: 2px solid red;');",
						query);
				//Thread.sleep(200);
				js.executeScript("arguments[0].setAttribute('style','border: solid 2px white')", query);
				js.executeScript("arguments[0].click();", query);
				setCell(workbook, row, "Pass");
			} else {
				setCell(workbook, row, "Fail");
				screenShot_failure(workbook, driver, row, value, dir);
			}
		} catch (IOException e) {
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	private static void pkgdisplayed(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			String value1, File dir) {
		try {

			int fibe = driver.findElements(By.xpath("(//label[contains(text(),'Fibe')])")).size();

			for (int i = 0; i < fibe; i++) {
				String fib = driver.findElement(By.xpath(xpathId)).getText();

				if (fib.equals(value)) {
					driver.findElement(By.xpath(value1)).click();
					setCell(workbook, row, "Pass");
					// break;
				} else {
					setCell(workbook, row, "Fail");
					screenShot_failure(workbook, driver, row, value, dir);
					break;
				}

			}
			// boolean isEnabled = driver.findElement(By.xpath(xpathId))
			// .isDisplayed();
			//
			//
			// if (isEnabled==true) {
			// setCell(workbook, row, "Pass");
			// } else {
			// setCell(workbook, row, "Fail");
			// screenShot_failure(workbook,driver, row, value,dir);
			// }
		} catch (IOException e) {
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * scroll down
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param xpathId
	 * @param value
	 * @throws IOException
	 */
	private static void scrollDown(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String xpathId, String value,
			File dir) throws IOException {
		try {

			Actions dragger = new Actions(driver);
			WebElement draggablePartOfScrollbar = driver.findElement(By.xpath(xpathId));

			// drag downwards
			int numberOfPixelsToDragTheScrollbarDown = Integer.parseInt(value);
			for (int i = 10; i < 500; i = i + numberOfPixelsToDragTheScrollbarDown) {
				try {
					// this causes a gradual drag of the scroll bar, 10 units at
					// a time
					dragger.moveToElement(draggablePartOfScrollbar).clickAndHold()
							.moveByOffset(0, numberOfPixelsToDragTheScrollbarDown).release().perform();
					Thread.sleep(1000L);

				} catch (Exception e1) {
				}
			}
			setCell(workbook, row, "Pass");

		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}

	/**
	 * zooming the screen
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void zoom(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value) throws IOException {
		try {
			String t = value;
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("document.body.style.zoom='" + t + "%'");
			Thread.sleep(3000);
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}
/*
 window authentication using AutoIT Code for SSO
 */
	private static void windowAuthentication(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir) throws IOException {
		try {
			Process p = Runtime.getRuntime().exec("C:\\SeleniumSetup\\Auto.exe");
			p.waitFor();
			Thread.sleep(1000);
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);

		}
	}


	/**
	 * Change screen resolution
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @param value
	 * @throws IOException
	 */
	private static void changeResolution(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value)
			throws IOException {
		try {
			String val1 = value.substring(0, value.indexOf(","));
			String val2 = value.substring(value.indexOf(","), value.length());
			driver.manage().window()
					.setSize(new org.openqa.selenium.Dimension(Integer.parseInt(val1), Integer.parseInt(val2)));
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * maximize the window
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @throws IOException
	 */
	private static void maximize(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {
			driver.manage().window().maximize();

			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
		}
	}

	/**
	 * maximize the window
	 *
	 * @param workbook
	 * @param row
	 * @param driver
	 * @throws IOException
	 */
	private static void CalculateTime(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, long Stime, long Etime)
			throws IOException {
		try {

			Properties props = new Properties();
			props.load(new FileInputStream(browserpath));
			// Thread.sleep(3000);
			row.createCell(7)
					.setCellValue(String.format("%d hr,%d min, %d sec, %d milli sec",
							TimeUnit.MILLISECONDS.toHours(Etime - Stime),
							TimeUnit.MILLISECONDS.toMinutes(Etime - Stime)
									- TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(Etime - Stime)), // The
																												// change
																												// is
																												// in
																												// this
																												// line
					TimeUnit.MILLISECONDS.toSeconds(Etime - Stime)
							- TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(Etime - Stime)),
					(Etime - Stime) - TimeUnit.SECONDS.toMillis(TimeUnit.MILLISECONDS.toSeconds(Etime - Stime))));

			// row.createCell(6).setCellValue(Etime - Stime);
			fileInputStream.close();
			//
			if (browsername.equals("firefox")) {
				// Thread.sleep(3000);
				outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));
			} else if (browsername.equals("iexplorer")) {

				outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
			} else if (browsername.equals("chrome")) {

				outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
			} else {
				outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
			}

			workbook.write(outfile);

			outfile.close();
		} catch (Exception e) {
			printErrorMessage(row, e);
			e.printStackTrace();

		}

	}

	private static void fileupload(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, File dir)
			throws IOException {
		try {
			Runtime.getRuntime().exec(value);
			setCell(workbook, row, "Pass");
		} catch (Exception e) {
			setCell(workbook, row, "Fail");
			screenShot_failure(workbook, driver, row, value, dir);
			e.printStackTrace();
			printErrorMessage(row, e);
			tobecontinue = false;
		}
	}

	/**
	 * check if excel file open
	 *
	 * @param serviceName
	 * @return
	 * @throws Exception
	 */
	public static boolean isProcessRunging(String serviceName) throws Exception {
		Process p = Runtime.getRuntime().exec("tasklist");
		BufferedReader reader = new BufferedReader(new InputStreamReader(p.getInputStream()));
		String line;
		while ((line = reader.readLine()) != null) {
			if (line.contains(serviceName)) {
				return true;
			}
		}
		return false;
	}

	public static void validatetable(HSSFWorkbook workbook, HSSFRow row, WebDriver driver, String value, String value1)
			throws IOException

	{
		int rc = driver.findElements(By.xpath(".//*[@id='fullListing-tpl']/div/table/tbody/tr")).size();

		System.out.println(rc);

		String first_part1 = "(.//*[@id='fullListing-tpl']/div/table/tbody/tr/td[1])[";
		String second_part1 = "]";
		String first_part2 = "(.//*[@id='fullListing-tpl']/div/table/tbody/tr/td[2])[";
		String second_part2 = "]";
		for (int j = 1; j <= rc; j++) { // String cmp1=value;
										// String cmp2=value1;
			String final_xpath_c1 = first_part1 + j + second_part1;
			String final_xpath_c2 = first_part2 + j + second_part2;

			System.out.println(final_xpath_c1);
			System.out.println(final_xpath_c2);
			String Table_data1 = driver.findElement(By.xpath(final_xpath_c1)).getText();
			System.out.println(Table_data1);
			String Table_data2 = driver.findElement(By.xpath(final_xpath_c2)).getText();
			System.out.println(Table_data2);
			if (value.equals(Table_data1)) {
				// String Table_data2 =
				// driver.findElement(By.xpath(final_xpath_c2)).getText();
				// System.out.println(Table_data2);
				if (value1.equals(Table_data2)) {
					System.out.println("Pass");
					setCell(workbook, row, "Pass");
					break;
				}

				else {
					System.out.println("Fail");
					setCell(workbook, row, "Fail");

					break;
				}
				// break;
			}
			// break;
		}
	}

	// --------------Report---------------------

	public void generateExcelReport(HSSFWorkbook workbook, HSSFSheet autoScript, HSSFSheet reportSheet,
			String testCaseName, HSSFSheet worksheet, String time_min1, String dat, String start_dat,
			Long finaltimeinsec) throws ParserConfigurationException, SAXException, IOException {
		System.out.println("Inside generate");

		numOfTestcases++;
		int rows = autoScript.getPhysicalNumberOfRows();
		System.out.println("Rows is :"+rows);
		int rows_master = worksheet.getPhysicalNumberOfRows();
		System.out.println("Rows is :"+rows_master);
		int j = 1;
		int i = 1;

		int fail = 0;
		int pass = 0;

		try {
			while (i < rows) {
				HSSFRow autoScriptRow = autoScript.getRow(i);
				if (testCaseName.trim().equals(autoScriptRow.getCell(0).toString().trim())) {
					// if(testCaseName.equals(autoScriptRow.getCell(0))){
					System.out.println(testCaseName);
					// if(testCaseName.equals(autoScriptRow.getCell(0).toString())){
					//isValid = true;
					if ("Pass".equals(autoScriptRow.getCell(6).toString().trim())) {
						// if("Pass".equals(autoScriptRow.getCell(7))){
						pass++;
						System.out.println(pass+" Pass");
					} else {
						System.out.println("Failed the step");
							fail++;
						break;
					}

				}
				i++;
			}
			System.out.println("Before print");
			System.out.println(fail+" Fail");
			System.out.println(pass+" Pass");
			HSSFRow row = reportSheet.createRow(numOfTestcases);
			CellStyle style = workbook.createCellStyle();
			style.setFillPattern(HSSFCellStyle.FINE_DOTS);
			style.setFillForegroundColor(HSSFColor.GREEN.index);
			// style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

			CellStyle style1 = workbook.createCellStyle();
			style1.setFillPattern(HSSFCellStyle.FINE_DOTS);
			style1.setFillForegroundColor(HSSFColor.RED.index);
			// style1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			style1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

			//if (isValid) {
				System.out.println(fail);
				row.createCell(0).setCellValue(testCaseName);

				if (fail > 0) {
					System.out.println("Test Case is failed");
					row.createCell(1).setCellValue("Fail");

					Cell cell = row.createCell(1);
					cell.setCellValue("Fail");
					cell.setCellStyle(style1);

				}

				else {
					System.out.println("Test Case is passed");
					row.createCell(1).setCellValue("Pass");
					Cell cell = row.createCell(1);
					cell.setCellValue("Pass");
					cell.setCellStyle(style);
				}

				row.createCell(2).setCellValue(fail + pass);
				row.createCell(3).setCellValue(pass);
				row.createCell(4).setCellValue(fail);

			//}
			while (j < rows_master) {
				row.createCell(5).setCellValue(time_min1);
				row.createCell(6).setCellValue(start_dat);
				// row.createCell(7).setCellValue(time_end1);
				row.createCell(7).setCellValue(dat);

				j++;

			}

		} catch (Exception e) {
			// setCell(workbook, row, "Fail");
			// e.printStackTrace();
			// printErrorMessage(row, e);
			System.out.println("Hello I am inside catch block");
		}
		try {

			System.out.println("reportsheet" + reportSheet.getSheetName());

			// String testCaseName = row.getCell(2).getStringCellValue();

			System.out.println("testcasename" + testCaseName);

			//
			Properties props = new Properties();
			props.load(new FileInputStream(browserpath));

			fileInputStream.close();
			if (browsername.equals("firefox")) {
				outfile = new FileOutputStream(props.getProperty("firefox.excelPath"));

			} else if (browsername.equals("iexplorer")) {
				outfile = new FileOutputStream(props.getProperty("iexplorer.excelPath"));
			} else if (browsername.equals("safari")) {
				outfile = new FileOutputStream(props.getProperty("safari.excelPath"));
			} else {
				outfile = new FileOutputStream(props.getProperty("chrome.excelPath"));
			}

			// System.out.println(outfile);
			workbook.write(outfile);
			System.out.println("successful updation in the sheet");
			outfile.close();
			// }
			//

		} catch (Exception e) {
			System.out.println("Hello Im outside the method block()");
		}


	}


}
