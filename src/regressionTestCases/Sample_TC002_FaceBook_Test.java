package regressionTestCases;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import configurationFiles.CONSTANTS;
import generalFunctions.WebApp_GenericFunctions;
import testUtilities.TestUtilities;

public class Sample_TC002_FaceBook_Test extends WebApp_GenericFunctions{
	
	/**
	 * @param ApplicationPageToLaunch 
	 * @param uname
	 * @throws Exception
	 */
	@Test(priority = 0, enabled = true, dataProvider = "testgetdata")
	public void TC001(String ApplicationPageToLaunch,
			String CustomerSearchCriteria, 
			String CustomerNumber)	throws Exception {
		BrowserSetUp();
	}
  
/**
 * @return
 */
@DataProvider(name="testgetdata")
	public Object[][] testgetdata() {
		try {
			Object[][] testgetdata = TestUtilities.getTestDataBasedOnTestCase(CONSTANTS.TCxls, this.getClass().getSimpleName());
			return testgetdata;
		} catch (Throwable e) {
			return null;
		}
	}
}
