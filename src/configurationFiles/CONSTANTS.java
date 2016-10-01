/**
 * 
 */
package configurationFiles;

import java.awt.Robot;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import testUtilities.Xls_ReadWrite;

/**
 * @author natarsa
 *
 */
public final class CONSTANTS {
	
	public static int StaticWaitTime=10;
	public static Robot robot;
	public static boolean acceptNextAlert = true;
	public static StringBuffer verificationErrors = new StringBuffer();
	public static Xls_ReadWrite TCxls=null;
	public static Xls_ReadWrite Credxls = null;
	public static String LDAPUsername = "";
	public static String LDAPPassword = "";
	public static String EnvironmentToRun = "";
	public static String EmailAddress = "";
	public static String EmailPassword = "";
	public static String DBUsername="";
	public static String DBPassword="";
	//To be used for future purpose
	public static String URLTobeUsed = "";
	public static String ServerName="serverName";
	public static String ServerPortNumber="50000";
	public static String DatabaseToConnect="DatabaseName";
	
	public static String fs=File.separator;
	public static final String Dateformat = new SimpleDateFormat("_dd_MM_yyyy_HH_mm_ss_SSS").format(new Date());
	public static final String ExtentConfigPath=System.getProperty("user.dir")+fs+"src"+fs+CONSTANTS.class.getPackage().getName().replace(".",fs)+fs+"extent-config.xml";
	public static String ScreenshotsPath=System.getProperty("user.dir")+fs+"ExtentReports"+fs+"ErrorScreenshots";
	public static final String strDownloadPath=System.getProperty("user.dir")+fs+"Downloads"+fs;
	public static final String IEDriverTemp=System.getProperty("user.dir")+fs+"IEDrivertemp"+fs;
	public static final String Uploads=System.getProperty("user.dir")+fs+"Uploads"+fs;
	public static final String chromeDriverPath=System.getProperty("user.dir")+fs+"src"+fs+CONSTANTS.class.getPackage().getName().replace(".",fs)+fs+"WebDriver_chromedriver.exe";
	public static final String IEDriverPath=System.getProperty("user.dir")+fs+"src"+fs+CONSTANTS.class.getPackage().getName().replace(".",fs)+fs+"WebDriver_IEDriverServer.exe";
	public static final String operaDriverPath=System.getProperty("user.dir")+fs+"src"+fs+CONSTANTS.class.getPackage().getName().replace(".",fs)+fs+"WebDriver_operadriver.exe";	
		
	/**
	 * Constants for Expected result
	 */
	}
