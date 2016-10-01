package configurationFiles;

import java.io.File;

public class Address_DataSheet_Constants {

	
	public static String fs=File.separator;
	/**
	 * Test Data Worksheet part for all applications
	 */
	public static String TESTCASES_WORKSHEET ="Testcases";
	public static String TESTCASES_TESTCASEID="TestCaseID";
	public static String TestData_WORKSHEET = "ApplicationLevelTestData";
	
	/*
	 * CREDENTIALS Worksheet
	 */
	public static String CRED_INPUT_FILE=System.getProperty("user.dir")+fs+"TestData"+fs+"ApplicationsHub.xlsx";
	public static String CRED_WORKSHEET ="Credentials";
	public static String CRED_LOGGEDUSER="CurrentlyLoggedUser_OR_AppUser";
	/*
	 * To run WAS Applications
	 */
	public static String APPTORUN_WORKSHEET ="RunApplications";
	public static String APP_NAMETOLOAD="ApplicationsToRun";
	public static String RUNMODE="Runmode";
	/**
	 * Captcha Worksheet 
	 */
	public static String CaptchaTestData_WORKSHEET = "CaptchaTestData"; 
	public static String CAPTCHA_TEXT="CaptchaText";
	public static String CaptchaTestData_CaptchaImage = "CaptchaImage";
	/**
	 * URL_Applications Worksheet
	 */
	public static String APPLICATIONSURLPARAMETERS_BASEDONENVSELCTED_WORKSHEET="ApplicationURLParameters";
	public static String APPLICATIONSURLPARAMETERS_ApplicationsToRunWRTEnvironment = "ApplicationsToRunWRTEnvironment"; 
	public static String APPLICATIONSURLPARAMETERS_VersionOfTheApplication = "VersionOfTheApplication";
	public static String APPLICATIONSURLPARAMETERS_URLTobeUsed = "URLTobeUsed";
	/*
	 * RUN MODES FOR BROWSERS
	 */
	public static String RUN_ON_FIREFOX = "RunOnFirefox";
	public static String RUN_ON_CHROME = "RunOnChrome";
	public static String RUN_ON_IE = "RunOnIE";
	public static String RUN_ON_SAFARI = "RunOnSafari";
	public static String RUN_ON_OPERA = "RunOnOpera";
	public static String RUN_ON_HEADLESS = "RunOnHeadless";
	public static String YES="Y";

}
