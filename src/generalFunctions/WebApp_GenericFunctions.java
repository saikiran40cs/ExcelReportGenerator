/*'*************************************************************************************************************************************************
' Script Name			: WebApp Functions
' Description			: Used to call the functions for all the web applications
' How to Use			:
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Author                    Version          Creation Date         Reviewer Name         Reviewed Date           Comments 
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Sai Kiran 		         v0.1             06-January-2016
'*************************************************************************************************************************************************
 */
package generalFunctions;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.TimeZone;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Proxy.ProxyType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.opera.OperaDriver;
import org.openqa.selenium.opera.OperaOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.safari.SafariOptions;
import org.testng.ITestResult;
import org.testng.Reporter;
import org.testng.SkipException;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.xml.sax.SAXException;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import configurationFiles.Address_DataSheet_Constants;
import configurationFiles.CONSTANTS;
import junit.framework.AssertionFailedError;
import testUtilities.ExtentManager;
import testUtilities.TestUtilities;
import testUtilities.Xls_ReadWrite;

public class WebApp_GenericFunctions {
	public WebDriver driver;
	public Properties OR = null;
	public String RunOnBrowser = null;
	public String baseURL = "";
	public String temp = null;
	public String CurrentRunningTCName;
	public String GetApplicationName;
	public ExtentReports extent;
	public ExtentTest test;
	public String ExecutionStartTime,ExecutionEndTime;
	public String getCurrentlyLoggedInUser;
	// Create an Report in excel
	XSSFWorkbook book = new XSSFWorkbook();
	String sheetName ="TestSuite Report";
	int rowNum;


	@BeforeSuite
	/**
	 * Function to change the Error Screenshot folder name before the Suite starts
	 * @author saikiran.nataraja
	 */
	public void CreateErrorRepFolder() throws ParserConfigurationException, SAXException, IOException{
		CONSTANTS.ScreenshotsPath=CONSTANTS.ScreenshotsPath+CONSTANTS.Dateformat+"ErrorScreenshots";
		initializeExcelReport();
	}

	/**
	 * Function to generate excel report based on the test scripts
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 * @throws IOException
	 */
	public void initializeExcelReport() throws ParserConfigurationException, SAXException, IOException {
		File reportFile = new File(Address_DataSheet_Constants.REP_INPUT_FILE); 
		if(reportFile.exists()){
			Reporter.log("reportFile exists");
		}else{

			XSSFSheet sheet,sheet1;
			FileOutputStream writeXLOutput = null;
			XSSFCellStyle TopHeaderContents = book.createCellStyle();
			XSSFCellStyle TableHeader = book.createCellStyle();
			XSSFCellStyle TableContentsOnLeft = book.createCellStyle();
			XSSFCellStyle TableContentsOnRight = book.createCellStyle();
			XSSFRow row;
			//Header Column Names
			XSSFCell cel_AppName;
			XSSFCell cel_name;
			XSSFCell cel_status;
			XSSFCell cel_exp;
			XSSFCell cel_ExecStartTime;
			XSSFCell cel_ExecEndTime;
			// Set the Font details for the entire sheet
			XSSFFont defaultFont,HeaderFont;
			String path = WebApp_GenericFunctions.class.getClassLoader().getResource("./").getPath();
			path = path.replace("bin", "src");
			rowNum=0;
			HeaderFont = book.createFont();
			HeaderFont.setFontHeightInPoints((short) 11);
			HeaderFont.setFontName("Calibri");
			HeaderFont.setColor(IndexedColors.WHITE.getIndex());
			HeaderFont.setBold(true);
			HeaderFont.setItalic(false);

			defaultFont = book.createFont();
			defaultFont.setFontHeightInPoints((short) 11);
			defaultFont.setFontName("Calibri");
			defaultFont.setColor(IndexedColors.BLACK.getIndex());
			defaultFont.setBold(true);
			defaultFont.setItalic(false);

			// create style for cells in header row
			TopHeaderContents.setFont(HeaderFont);
			TopHeaderContents.setFillPattern(XSSFCellStyle.NO_FILL);
			TopHeaderContents.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
			TopHeaderContents.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);		
			// Set the border style for the workbook
			TopHeaderContents.setAlignment(XSSFCellStyle.ALIGN_LEFT);

			// create style for cells in header row
			TableHeader.setFont(defaultFont);
			TableHeader.setFillPattern(XSSFCellStyle.NO_FILL);
			TableHeader.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
			TableHeader.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
			// Set the border style for the workbook
			TableHeader.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			TableHeader.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			TableHeader.setBorderRight(XSSFCellStyle.BORDER_THIN);
			TableHeader.setBorderTop(XSSFCellStyle.BORDER_THIN);
			TableHeader.setAlignment(XSSFCellStyle.ALIGN_LEFT);


			// create style for cells in table contents
			TableContentsOnRight.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			TableContentsOnRight.setFillPattern(XSSFCellStyle.NO_FILL);
			TableContentsOnRight.setWrapText(false);
			// Set the border style for the workbook
			TableContentsOnRight.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			TableContentsOnRight.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			TableContentsOnRight.setBorderRight(XSSFCellStyle.BORDER_THIN);
			TableContentsOnRight.setBorderTop(XSSFCellStyle.BORDER_THIN);
			TableContentsOnRight.setAlignment(XSSFCellStyle.ALIGN_RIGHT);

			// create style for cells in table contents
			TableContentsOnLeft.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			TableContentsOnLeft.setFillPattern(XSSFCellStyle.NO_FILL);
			//		TableContentsOnLeft.setWrapText(false);
			// Set the border style for the workbook
			TableContentsOnLeft.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderRight(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderTop(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setAlignment(XSSFCellStyle.ALIGN_LEFT);

			//Create worksheet
			sheet1 = book.createSheet("Graphical_Summary");
			row = sheet1.createRow(1);
			XSSFCellStyle hiddenstyle = book.createCellStyle();
			hiddenstyle.setHidden(true);

			//Create worksheet
			sheet = book.createSheet(sheetName);
			book.setActiveSheet(1);
			row = sheet.createRow(rowNum++);
			//Header Column Names
			cel_AppName = row.createCell(0);
			cel_AppName.setCellValue("Automated Test Execution Report");
			cel_AppName.setCellStyle(TopHeaderContents);

			row = sheet.createRow(rowNum++);
			cel_AppName = row.createCell(0);
			cel_AppName.setCellValue("");

			row = sheet.createRow(rowNum++);
			cel_name = row.createCell(1);
			cel_status = row.createCell(2);
			cel_name.setCellValue("Number of Test cases passed");
			cel_status.setCellFormula("COUNTIFS(C10:C103,\"PASSED\")");
			cel_name.setCellStyle(TableContentsOnLeft);
			cel_status.setCellStyle(TableContentsOnRight);

			row = sheet.createRow(rowNum++);
			cel_name = row.createCell(1);
			cel_status = row.createCell(2);
			cel_name.setCellValue("Number of Test cases failed");
			cel_status.setCellFormula("COUNTIFS(C10:C103,\"FAILED\")");
			cel_name.setCellStyle(TableContentsOnLeft);
			cel_status.setCellStyle(TableContentsOnRight);

			row = sheet.createRow(rowNum++);
			cel_name = row.createCell(1);
			cel_status = row.createCell(2);
			cel_name.setCellValue("Number of Test cases Not Executed");
			cel_status.setCellFormula("COUNTIFS(C10:C103,\"NOT EXECUTED\")");
			cel_name.setCellStyle(TableContentsOnLeft);
			cel_status.setCellStyle(TableContentsOnRight);

			row = sheet.createRow(rowNum++);
			cel_name = row.createCell(1);
			cel_status = row.createCell(2);
			cel_name.setCellValue("TOTAL");
			cel_status.setCellFormula("SUM(C3:C5)");
			cel_name.setCellStyle(TableContentsOnLeft);
			cel_status.setCellStyle(TableContentsOnRight);

			row = sheet.createRow(rowNum++);
			cel_AppName = row.createCell(0);
			cel_AppName.setCellValue("");

			row = sheet.createRow(rowNum++);
			//Header Column Names
			cel_AppName = row.createCell(0);
			cel_AppName.setCellValue("Test case level execution details");
			cel_AppName.setCellStyle(TopHeaderContents);

			row = sheet.createRow(rowNum++);
			//Header Column Names
			cel_AppName = row.createCell(0);
			cel_name = row.createCell(1);
			cel_status = row.createCell(2);
			cel_ExecStartTime=row.createCell(3);
			cel_ExecEndTime=row.createCell(4);
			cel_exp = row.createCell(5);

			// Set the Header names for the sheet
			cel_AppName.setCellValue("Application Name");
			cel_AppName.setCellStyle(TableHeader);
			cel_name.setCellValue("Test Script Name");
			cel_status.setCellValue("Execution Status");
			cel_ExecStartTime.setCellValue("Execution Start Time");
			cel_ExecEndTime.setCellValue("Execution End Time");
			cel_exp.setCellValue("Comments");
			cel_exp.setCellStyle(TableHeader);
			// set style for the table header
			cel_name.setCellStyle(TableHeader);
			cel_status.setCellStyle(TableHeader);
			cel_ExecStartTime.setCellStyle(TableHeader);
			cel_ExecEndTime.setCellStyle(TableHeader);
			// Auto size the column widths based on the names
			sheet.autoSizeColumn(0);
			sheet.autoSizeColumn(1);
			sheet.autoSizeColumn(2);
			sheet.autoSizeColumn(3);
			sheet.autoSizeColumn(4);
			sheet.autoSizeColumn(5);

			writeXLOutput = new FileOutputStream(Address_DataSheet_Constants.REP_INPUT_FILE);
			book.write(writeXLOutput);
			writeXLOutput.close();
			book.close();
		}
		CONSTANTS.Reportxls=new Xls_ReadWrite(Address_DataSheet_Constants.REP_INPUT_FILE);
	}

	/**
	 * Function to Setup the Call to other business functions to others. Always call the instantiating class from the other.
	 * @author saikiran.nataraja
	 * @param DriverOfSending
	 * @param ExtReporttest
	 * @param BrowserRunningOn
	 * @param ApplicationName
	 */
	public void SetupToCallOtherBusinessFunctionsFromOther(WebDriver DriverOfSending, ExtentTest ExtReporttest,String BrowserRunningOn,String ApplicationName) {
		Properties ReqOR = null;
		try {
			ReqOR = ObjectProperty_Setup(ReqOR,ApplicationName);
			OR=ReqOR;
			if(DriverOfSending != null){
				driver=DriverOfSending;
				test=ExtReporttest;
				RunOnBrowser=BrowserRunningOn;
				GetApplicationName=ApplicationName;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Function to read the properties File , Loading the values to the Property Variable
	 * @author saikiran.nataraja
	 * @param ObjpropName
	 * @param prFileLoad
	 * @return ObjProperties for which it is called
	 * @throws Exception
	 */
	public  Properties ObjectProperty_Setup(Properties ObjpropName ,String prFileLoad) throws Exception {
		FileInputStream fi = null;
		try {
			prFileLoad=System.getProperty("user.dir")+CONSTANTS.fs+"src"+CONSTANTS.fs+WebApp_GenericFunctions.class.getCanonicalName().replace("generalFunctions.WebApp_GenericFunctions", "objectMaps").replace(".", CONSTANTS.fs)+CONSTANTS.fs+prFileLoad+"Object.properties";
			//			System.out.println(prFileLoad);
			ObjpropName = new Properties();
			fi = new FileInputStream(prFileLoad);
			ObjpropName.load(fi);
			fi.close();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (NullPointerException e) {
			e.printStackTrace();
			//			System.out.println("Error in Loading ASJ Property file"+e.getMessage());
		}
		return ObjpropName;
	}

	/**
	 * To Check Test case class is runnable or not is checked here.
	 * @author saikiran.nataraja
	 * @param TCxls_INPUTFILE
	 * @param callerClassName
	 * @param Worksheet_Name
	 * @param Worksheet_FieldNameToSearch
	 */
	public void CheckRunMode(Xls_ReadWrite TCxls_INPUTFILE,String callerClassName,String Worksheet_Name,String Worksheet_FieldNameToSearch) {

		System.out.println(callerClassName);
		if(!TestUtilities.isTestCaseRunnable(TCxls_INPUTFILE,callerClassName,Worksheet_Name,Worksheet_FieldNameToSearch)){
			throw new SkipException("Skipping Test Case \'"+callerClassName+"\' as runmode set to NO");
		}
		else if (!TestUtilities.RunTestcaseOnBrowser(TCxls_INPUTFILE,callerClassName, RunOnBrowser)){
			throw new SkipException("Skipping Test Case \'"+callerClassName+"\' as run on \'"+RunOnBrowser+ "\' browser set to NO");
		}
	}

	/**
	 * This function is used to get the credentials from the Applications HUb Worksheet
	 * @author saikiran.nataraja
	 * @param Credxls
	 * @param ApplicationLDAPUser
	 */
	public  void getCredentials(Xls_ReadWrite Credxls,String... ApplicationLDAPUser){
		if (!Credxls.isSheetExist(Address_DataSheet_Constants.CRED_WORKSHEET)) {
			Credxls = null;
			throw new SkipException("The test data worksheet is not availble for test case "+Address_DataSheet_Constants.CRED_WORKSHEET+". HEnce skipped the test caseexecution.");
		}
		int cols = Credxls.getColumnCount(Address_DataSheet_Constants.CRED_WORKSHEET);
		//Always read the Environment details of the currently executing user
		for (int i = 2; i <= Credxls.getRowCount(Address_DataSheet_Constants.CRED_WORKSHEET); i++) {
			if (Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, Address_DataSheet_Constants.CRED_LOGGEDUSER, i).equalsIgnoreCase(getCurrentlyLoggedInUser)) {
				for (int colNum = 0; colNum < cols; colNum++) {
					if (Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, "LDAPUsername", i).equals("@null")){
						//	System.out.println("The column name \'LDAPUsername\' is not available in test data excel sheet");
					}
					CONSTANTS.EnvironmentToRun =Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, "EnvironmentToRun", i).toString().replace("_", "");
				}break;
			}
		}
		//Below code is a special case with respect to the applications User Credentials.. For Example we use different Credentials for accessing YVP Application via LDAP.
		if(ApplicationLDAPUser.length>0){
			getCurrentlyLoggedInUser=ApplicationLDAPUser[0];
		}
		for (int i = 2; i <= Credxls.getRowCount(Address_DataSheet_Constants.CRED_WORKSHEET); i++) {
			if (Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, Address_DataSheet_Constants.CRED_LOGGEDUSER, i).equalsIgnoreCase(getCurrentlyLoggedInUser)) {
				for (int colNum = 0; colNum < cols; colNum++) {
					if (Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, "LDAPUsername", i).equals("@null")){
						//	System.out.println("The column name \'LDAPUsername\' is not available in test data excel sheet");
					}
					CONSTANTS.LDAPUsername = Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, "LDAPUsername", i);
					CONSTANTS.LDAPPassword =Credxls.getCellData(Address_DataSheet_Constants.CRED_WORKSHEET, "LDAPPassword", i);
				}break;
			}
		}
		Credxls=null;
	}

	/**
	 * Function to get the URL Parameters from the Applications Hub Excel sheet and get the URL based on the EnvironmentToRun
	 * which is available in "Credentials" worksheet 
	 * @author saikiran.nataraja
	 * @param ApplicationName
	 * @return 
	 */
	public String getURLParameters(String ApplicationName){
		Xls_ReadWrite AppURLxls = new Xls_ReadWrite(Address_DataSheet_Constants.CRED_INPUT_FILE);
		if (!AppURLxls.isSheetExist(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET)) {
			AppURLxls = null;
			throw new SkipException("The test data worksheet is not availble for test case "
					+ Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET + ". Hence skipped the test case execution.");
		}
		int cols = AppURLxls.getColumnCount(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET);
		for (int i = 2; i <= AppURLxls.getRowCount(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET); i++) {
			if (AppURLxls.getCellData(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET,
					Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_ApplicationsToRunWRTEnvironment, i).equalsIgnoreCase(ApplicationName)) {
				for (int colNum = 0; colNum < cols; colNum++) {
					if (AppURLxls.getCellData(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET,"URLTobeUsed", i).equals("@null")) {
						assertEquals(true, "Unable to Find the URL to be used from Applications HUB worksheet","");
					}
					baseURL = AppURLxls.getCellData(Address_DataSheet_Constants.APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET,
							"URLTobeUsed", i);
				}
				break;
			}
		}
		AppURLxls=null;
		return baseURL;
	}

	/**
	 *Setup The Object Map before Running Scripts
	 *@author saikiran.nataraja
	 *@throws Exception
	 */
	@Parameters("BrowserType")
	@BeforeTest
	public void setUp(@Optional("firefox") String Browser) throws Exception {
		try{
			//Creating Xls_ReadWrite object.
			CurrentRunningTCName=super.getClass().getSimpleName();
			getCurrentlyLoggedInUser=System.getProperty("user.name");
			CONSTANTS.sdf.setTimeZone(TimeZone.getTimeZone("EET"));
			ExecutionStartTime=CONSTANTS.sdf.format(new Date());
			GetApplicationName=CurrentRunningTCName.substring(0, CurrentRunningTCName.indexOf("_"));
			RunOnBrowser=Browser;
			//Load the ApplicationsHub Workbook into the credxls object
			CONSTANTS.Credxls=new Xls_ReadWrite(Address_DataSheet_Constants.CRED_INPUT_FILE);
			//Choose Application URL
			switch (GetApplicationName) {
			case "Google":
				getCredentials(CONSTANTS.Credxls);
				baseURL=getURLParameters(GetApplicationName);
				break;
			default:
				baseURL = "NOT AVAILABLE";
			}
			//Application specific test data load
			String ExcelFileToLoad = System.getProperty("user.dir")+CONSTANTS.fs+"TestData"+CONSTANTS.fs+GetApplicationName+"_Testdata.xlsx";
			CONSTANTS.TCxls= new Xls_ReadWrite(ExcelFileToLoad);
			OR=ObjectProperty_Setup(OR,"WebApp");
		}
		catch(IOException e){
			e.printStackTrace();
		}	
		catch(NullPointerException e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * Function to intialise Extent report
	 */
	public void InitializeExtentReport(){
		//Instantiating the ExtentReports
		extent=ExtentManager.getInstance(CONSTANTS.EnvironmentToRun);
		test=extent.startTest(CurrentRunningTCName, "'"+CurrentRunningTCName+"' used to check details in "+GetApplicationName+" Application." );
		test.assignAuthor(getCurrentlyLoggedInUser);
		test.assignCategory("RegressionTestCases_"+GetApplicationName);
	}

	/**
	 * 	BrowserSetUp is an function to setup which browser webdriver should be enabled
	 * @author saikiran.nataraja
	 * @param Browser
	 * @param baseUrl
	 * @throws Exception
	 */ 
	public void BrowserSetUp() throws Exception {
		InitializeExtentReport();
		//Check the ApplicationHUB sheet is skippable
		if(!TestUtilities.isTestCaseRunnable(CONSTANTS.Credxls,GetApplicationName,Address_DataSheet_Constants.APPTORUN_WORKSHEET,Address_DataSheet_Constants.APP_NAMETOLOAD)){
			throw new SkipException("Skipping Test Case \'"+GetApplicationName+"\' as runmode set to NO");
		}
		//Checking Run Mode from the Respective application sheet
		CheckRunMode(CONSTANTS.TCxls,CurrentRunningTCName, Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.TESTCASES_TESTCASEID);
		CONSTANTS.robot = new Robot();
		String URLToLoad=baseURL;
		try{
			switch (RunOnBrowser) {
			case "firefox": 
				FirefoxProfile fp = new FirefoxProfile();					
				fp.setAcceptUntrustedCertificates(true); 	// Set profile to accept untrusted certificates
				fp.setAssumeUntrustedCertificateIssuer(false);	// Set profile to not assume certificate issuer is untrusted
				fp.setPreference("pdfjs.disabled", true);
				fp.setPreference("browser.download.folderList", 2);		//0- Desktop, 1-Browser's Default Path , 2- Custom Download Path
				fp.setPreference("plugin.scan.plid.all", false);
				fp.setPreference("plugin.scan.Acrobat", "99");
				fp.setPreference("browser.download.dir", CONSTANTS.strDownloadPath);
				fp.setPreference("browser.download.manager.alertOnEXEOpen", false);
				fp.setPreference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream;application/pdf");
				//					"application/msword,application/csv,text/csv,image/png ,image/jpeg, text/html,text/plain,application/octet-stream");
				fp.setPreference("browser.download.manager.showWhenStarting", false);
				fp.setPreference("browser.download.manager.focusWhenStarting", false);
				// fp.setPreference("browser.download.useDownloadDir", true);
				fp.setPreference("browser.helperApps.alwaysAsk.force", false);
				fp.setPreference("browser.download.manager.alertOnEXEOpen", false);
				fp.setPreference("browser.download.manager.closeWhenDone", false);
				fp.setPreference("browser.download.manager.showAlertOnComplete", false);
				fp.setPreference("browser.download.manager.useWindow", false);
				fp.setPreference("browser.download.manager.showWhenStarting", false);
				fp.setPreference("network.proxy.type", ProxyType.AUTODETECT.ordinal());
				fp.setPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
				fp.setPreference("plugin.disable_full_page_plugin_for_types", "application/pdf");	//,application/vnd.adobe.xfdf,application/vnd.fdf,application/vnd.adobe.xdp+xml
				driver = new FirefoxDriver(fp);
				break;
			case "chrome":
				ChromeOptions chromeOptions = new ChromeOptions();
				DesiredCapabilities ChromeCapabilities = DesiredCapabilities.chrome();
				ChromeCapabilities.setCapability(ChromeOptions.CAPABILITY, chromeOptions);					
				ChromeCapabilities.setCapability("network.proxy.type", ProxyType.AUTODETECT.ordinal());
				// Set ACCEPT_SSL_CERTS  variable to true
				ChromeCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
				ChromeCapabilities.setCapability(CapabilityType.ForSeleniumServer.ENSURING_CLEAN_SESSION, true); //17-05-2016 - Newly added to remove cache issues
				Map<String, Object> prefs = new HashMap<String, Object>();
				prefs.put("profile.default_content_settings.popups", 0);
				prefs.put("download.extensions_to_open", "pdf");
				prefs.put("download.prompt_for_download", "true");
				prefs.put("download.default_directory", CONSTANTS.strDownloadPath);
				chromeOptions.setExperimentalOption("prefs", prefs);
				System.setProperty("webdriver.chrome.driver",CONSTANTS.chromeDriverPath);
				System.setProperty("webdriver.chrome.args", "--disable-logging");
				System.setProperty("webdriver.chrome.silentOutput", "true");
				driver = new ChromeDriver(ChromeCapabilities);
				break;
			case "ie":
				DesiredCapabilities ieCapabilities = DesiredCapabilities.internetExplorer();
				// Set ACCEPT_SSL_CERTS  variable to true
				ieCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
				ieCapabilities.setCapability("network.proxy.type", ProxyType.AUTODETECT.ordinal());
				ieCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true); //Added to avoid the Protected Mode settings for all zones
				ieCapabilities.setCapability(InternetExplorerDriver.INITIAL_BROWSER_URL, "about:blank");
				ieCapabilities.setCapability(InternetExplorerDriver.SILENT, true);
				ieCapabilities.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION, true);
				ieCapabilities.setCapability(InternetExplorerDriver.IGNORE_ZOOM_SETTING,true);					
				System.setProperty("webdriver.ie.driver",CONSTANTS.IEDriverPath);
				//Add Pop-up allowed in IE manually
				driver = new InternetExplorerDriver(ieCapabilities);
				break;
				// Needs to be checked for Safari
			case "safari":
				SafariOptions safariOptions=new SafariOptions();
				DesiredCapabilities SafariCapabilities = DesiredCapabilities.safari();
				SafariCapabilities.setCapability(SafariOptions.CAPABILITY, safariOptions);
				// Set ACCEPT_SSL_CERTS  variable to true
				SafariCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
				SafariCapabilities.setCapability("network.proxy.type", ProxyType.AUTODETECT.ordinal());
				driver = new SafariDriver(SafariCapabilities);
				break;
			case "opera":
				OperaOptions operaOptions = new OperaOptions();					
				DesiredCapabilities OperaCapabilities = DesiredCapabilities.chrome();
				OperaCapabilities.setCapability(ChromeOptions.CAPABILITY, operaOptions);
				OperaCapabilities.setCapability("network.proxy.type", ProxyType.AUTODETECT.ordinal());
				// Set ACCEPT_SSL_CERTS  variable to true
				OperaCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
				System.setProperty("webdriver.opera.driver",CONSTANTS.operaDriverPath);
				System.setProperty("opera.arguments", "--disable-logging");
				System.setProperty("webdriver.opera.silentOutput", "true");
				driver = new OperaDriver(OperaCapabilities);
				break;
			default:
				Reporter.log("Driver Not Found");
				break;
			}
			driver.manage().window().maximize();
			//	driver.manage().deleteAllCookies();
			driver.get(URLToLoad);
			driver.manage().timeouts().implicitlyWait(3, TimeUnit.SECONDS);
			driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
			assertEquals(true, true, "Application URL: '"+baseURL+"' is opened in "+getBrowserVersion()+" successfully");
			Thread.sleep(CONSTANTS.StaticWaitTime);
		}
		catch(Exception e){
			e.printStackTrace();
		}

	}

	/**
	 * Function to return browser version 
	 * @return browser name and version
	 */
	public String getBrowserVersion() {
		String browser_version = null;
		Capabilities cap = ((RemoteWebDriver) driver).getCapabilities();
		String browsername = cap.getBrowserName();
		// This block to find out IE Version number
		if ("internet explorer".equalsIgnoreCase(browsername)) {
			boolean GetIEVersion=((String) ((JavascriptExecutor) driver).executeScript("return navigator.userAgent;")).contains("Trident/7.0");
			if(GetIEVersion){
				browser_version = "11.0";
			}
		} else
		{
			//Browser version for Firefox, Chrome and Opera
			browser_version = cap.getVersion();// .split(".")[0];
		}
		String browserversion = browser_version.substring(0, browser_version.indexOf("."));
		return Captialize(browsername) + " browser (Version: " + browserversion +" )";
	}

	/**
	 * Function to detect the OS in which it is running on
	 * @author saikiran.nataraja
	 * @returns the Operating system on which it was running on
	 */
	public static String OSDetector() {
		String os = System.getProperty("os.name").toLowerCase();
		if (os.contains("win")) {
			return "Windows";
		} else if (os.contains("nux") || os.contains("nix")) {
			return "Linux";
		} else if (os.contains("mac")) {
			return "Mac";
		} else if (os.contains("sunos")) {
			return "Solaris";
		} else {
			return "Other";
		}
	}

	/**
	 * Function to Captialize the word
	 * @param RequiredWord
	 * @return
	 */
	public String Captialize(String RequiredWord)
	{
		return RequiredWord.substring(0,1).toUpperCase()+RequiredWord.substring(1, RequiredWord.length()).toLowerCase();
	}

	/**
	 * AssertEquals function is used to verify the expected and actual value
	 * @author saikiran.nataraja
	 * @param expected value needs to be passed
	 * @param actual value retreived from the application
	 * @param message to be logged if the match is true
	 */
	public  void assertEquals(Object expected, Object actual,String message) {
		if (expected == null && actual == null) {
			test.log(LogStatus.FAIL, "Expected and Actual values are NULL");
			Reporter.log("Expected and Actual values are NULL", false);
			//			ExtentManager.extent.setTestRunnerOutput("Expected and Actual values are NULL"); 
			return;
		}
		if (expected != null && expected.equals(actual)) {
			test.log(LogStatus.PASS, message);
			Reporter.log(message);
			//			ExtentManager.extent.setTestRunnerOutput(message); 
			return;
		} else if (!expected.equals(actual)) {
			Reporter.log("Expected Result: '" + expected + "' does NOT match with Actual Result: '" + actual + "' .");
			extent.setTestRunnerOutput("Expected Result: '" + expected + "' does NOT match with Actual Result: '" + actual + "' .");
			Reporter.log("Error Snapshot is Located in : " + CONSTANTS.ScreenshotsPath + GetApplicationName + CONSTANTS.fs+ Captialize(RunOnBrowser) + "_" + CurrentRunningTCName + CONSTANTS.Dateformat + ".jpg");
			fail(format("", expected, actual));
		}
	}

	/**
	 * Sub to check Assertion failure messages
	 * @author saikiran.nataraja
	 * @param message
	 */
	public void fail(String message) {
		if (message == null) {
			throw new AssertionFailedError();
		}
		throw new AssertionFailedError(message);
	}

	/**
	 * Sub to report the issue in the required formatted text
	 * @author saikiran.nataraja
	 * @param message
	 * @param expected
	 * @param actual
	 * @return
	 */
	public  String format(String message, Object expected, Object actual) {
		String formatted = "The Expected Outcome is not matched with Actual Outcome::";
		if (message != null && message.length() > 0) {
			formatted = message + "  ";
		}
		return formatted + "Expected :'" + expected + "' but Actual is:'" + actual + "'";
	}

	/**
	 * Function to check element is present or not
	 * @author saikiran.nataraja
	 * @param xpathOfElement
	 * @return true - if xpath exists, false - if xpath doesnot exists.
	 */
	public boolean isXpathExists(String xpathOfElement) {
		try {
			driver.findElement(By.xpath(xpathOfElement));
			return true;
		} catch (Exception e) {
			return false;
		}
	}	    		         

	/**
	/**
	 * Capture the screenshot on error
	 * @param testResult
	 * @author saikiran.nataraja
	 * @throws Exception
	 */
	@AfterMethod 
	public void takeScreenShotOnFailure(ITestResult testResult) throws Exception {
		rowNum = CONSTANTS.Reportxls.findRowNumber(CurrentRunningTCName);
		//Write to the Excel only if the test case name, Application Name does NOT exist in the excel
		if(rowNum==0){
			rowNum = CONSTANTS.Reportxls.getRowCount(sheetName);
			CONSTANTS.Reportxls.setCellData(sheetName, "Test Script Name", rowNum, CurrentRunningTCName);
			CONSTANTS.Reportxls.setCellData(sheetName, "Application Name", rowNum, GetApplicationName);
		}
		//If the test name already exists then assign the row at which test case name exists to row number
		CONSTANTS.Reportxls.setCellData(sheetName, "Execution Start Time", rowNum, ExecutionStartTime);
		if (testResult.getStatus() == ITestResult.FAILURE){ 
			try{
				System.out.println(" - FAILED.");
				//Create Error Screenshot Directory if doesnot exists
				File dir = new File(CONSTANTS.ScreenshotsPath+CONSTANTS.fs+GetApplicationName+CONSTANTS.fs);
				dir.setWritable(true); //If SecurityManager.checkWrite(java.lang.String) method denies write access to the file.Hence made the directory writable
				dir.mkdirs();
				BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
				//Here the screenshot path is reduced to a maximum of 20 literals
				File imagePath=new File(CONSTANTS.ScreenshotsPath+CONSTANTS.fs+GetApplicationName+CONSTANTS.fs+CONSTANTS.Dateformat+Captialize(RunOnBrowser)+"_"+CurrentRunningTCName.substring(0, Math.min(CurrentRunningTCName.length(), 20)) +".jpg");
				imagePath.setWritable(true);
				ImageIO.write(image, "JPG", imagePath);
				CONSTANTS.Reportxls.setCellData(sheetName, "Execution Status", rowNum, "FAILED");
				CONSTANTS.Reportxls.setCellData(sheetName, "Comments", rowNum, testResult.getThrowable().getMessage());
				//Extent Reports take screenshot
				test.log(LogStatus.FAIL, "Failure Stack Trace: "+ testResult.getThrowable().getMessage());
				test.log(LogStatus.FAIL,"Snapshot below: " + test.addScreenCapture(imagePath.getAbsoluteFile().toString().replace(System.getProperty("user.dir")+CONSTANTS.fs+"ExtentReports"+CONSTANTS.fs, "")));			
			}    
			catch (Exception e){
				//			System.out.println("Class Utils | Method takeScreenshot | Exception occured while capturing ScreenShot : "+e.getMessage());
				//			e.printStackTrace();
			}
		}else if (testResult.getStatus() == ITestResult.SKIP) {
			CONSTANTS.Reportxls.setCellData(sheetName, "Execution Status", rowNum, "NOT EXECUTED");
			CONSTANTS.Reportxls.setCellData(sheetName, "Comments", rowNum, "Test skipped: " + testResult.getThrowable().getMessage());
			System.out.println(" - SKIPPED.");
			test.log(LogStatus.SKIP, "Test skipped: " + testResult.getThrowable().getMessage());
		}else{
			System.out.println(" - PASSED.");
			CONSTANTS.Reportxls.setCellData(sheetName, "Execution Status", rowNum, "PASSED");
			CONSTANTS.Reportxls.setCellData(sheetName, "Comments", rowNum, "");
			//Test is PASSED/INFO/
		}
		ExecutionEndTime=CONSTANTS.sdf.format(new Date());
		CONSTANTS.Reportxls.setCellData(sheetName, "Execution End Time", rowNum, ExecutionEndTime);
	}

	/**
	 * Sub to close the alert and get its text
	 * @return
	 */
	public String closeAlertAndGetItsText() {
		try {
			Alert alert = driver.switchTo().alert();
			String alertText = alert.getText();
			if (CONSTANTS.acceptNextAlert) {
				alert.accept();
			} else {
				alert.dismiss();
			}
			return alertText;
		} finally {
			CONSTANTS.acceptNextAlert = true;
		}
	}

	/**
	 * @return The calculated date required 
	 * @param ValidDateFormat - Valid Date format could be for eg: dd.MM.yyyy,..etc
	 * @param dateRequired - dateRequired field will take 0 for current date, positive number(1,2,..) for future date and negative number for Past dates 
	 */
	@SuppressWarnings("static-access")
	public String ProvidePastCurrentOrFutureDate(String ValidDateFormat, int dateRequired){

		//Declarations
		Date today = new Date();
		DateFormat dateformat = new SimpleDateFormat(ValidDateFormat); 	//Valid Date format could be for eg: dd.MM.yyyy,..etc
		Calendar cal = new GregorianCalendar();	 //Create a calender class to be instanciated

		//Calculating date
		cal.setTime(today);
		cal.add(cal.DAY_OF_MONTH, dateRequired);	//dateRequired field will take 0 for current date, positive number(1,2,..) for future date and negative number for Past dates 
		Date Past_Date1 = cal.getTime();

		//Return the calculated date
		return dateformat.format(Past_Date1);
	}

	/**
	 * Function to flash the element in webdriver
	 * @param element
	 * @param driver
	 */
	public static void flash(WebElement element, WebDriver driver) {
		JavascriptExecutor js = ((JavascriptExecutor) driver);
		String bgcolor  = element.getCssValue("backgroundColor");
		for (int i = 0; i <  3; i++) {
			changeColor("rgb(0,200,0)", element, js);
			changeColor(bgcolor, element, js);
			changeColor("rgb(0,200,0)", element, js);
		}
	}
	/**
	 * Function is internally called by flash function to highlight element using javascript
	 * @param color
	 * @param element
	 * @param js
	 */
	public static void changeColor(String color, WebElement element,  JavascriptExecutor js) {
		js.executeScript("arguments[0].style.backgroundColor = '"+color+"'",  element);
		try {
			Thread.sleep(20);
		}  catch (InterruptedException e) {
		}
	}

	/**
	 * Tears down the browser driver created.
	 * @author saikiran.nataraja
	 * @throws Exception
	 */
	@AfterClass
	public void tearDownTheClass()  {
		try {
			//End the Test
			extent.endTest(test);
			// write all resources to report file
			extent.flush();
			CONSTANTS.TCxls=null;
			OR=null;
			driver.close();
		} catch (Exception e) {
			//			 e.printStackTrace(); //remove comment while trying to debug the tests
		}
	}

	/**
	 * Used to Kill the remains of the Webdrivers in the processes after the suite execution
	 */
	@AfterSuite
	public void tearSuite(){
		try {
			CONSTANTS.Reportxls = null;
			Runtime rt = Runtime.getRuntime();
			rt.exec("taskkill /F /IM WebDriver_IEDriverServer.exe");
			rt.exec("taskkill /F /IM WebDriver_chromedriver.exe");
		} catch (IOException e) {
			//			e.printStackTrace(); //remove comment while trying to debug the tests
		}

	}
}