package testUtilities;

import org.testng.SkipException;

import configurationFiles.Address_DataSheet_Constants;

public class TestUtilities {

	// returns true if runmode of the testcase is set to Y
	/**
	 * @param xls
	 * @param testCaseName
	 * @return
	 */
	public static boolean isTestCaseRunnable(Xls_ReadWrite xls, String testCaseName,String TESTCASES_WORKSHEET,String TESTCASES_TESTCASEID){
		boolean isExecutable=false;
		for(int i=2; i<= xls.getRowCount(TESTCASES_WORKSHEET) ; i++){
			//String tcid=xls.getCellData("Test Cases", "TestcaseID", i);
			//String runmode=xls.getCellData("Test Cases", "Runmode", i);
			//System.out.println(testcaseid +" -- "+ runmode);
			
			if(xls.getCellData(TESTCASES_WORKSHEET, TESTCASES_TESTCASEID, i).equalsIgnoreCase(testCaseName)){
				if(xls.getCellData(TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUNMODE, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
					isExecutable= true;
					
				}else{
					isExecutable= false;
				}
			}//else
			//throw new SkipException("The selenium script is not availble to execute the test case "+testCaseName+ " menioned in testcase xlsx file. Hence skipped thae test case execution");
		}
		
		return isExecutable;
     }
	

	// returns true if run on browser ( firefox, ie, chrome)set to yes
	/**
	 * @param xls
	 * @param testCaseName
	 * @param browserSelection
	 * @return
	 */
	public static boolean  RunTestcaseOnBrowser(Xls_ReadWrite xls, String testCaseName, String browserSelection){
		boolean selectedBrowser=false;
		for(int i=2; i<= xls.getRowCount(Address_DataSheet_Constants.TESTCASES_WORKSHEET) ; i++){
			//String tcid=xls.getCellData("Test Cases", "TestcaseID", i);
			//String runmode=xls.getCellData("Test Cases", "Runmode", i);
			//System.out.println(testcaseid +" -- "+ runmode);

			if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.TESTCASES_TESTCASEID, i).equalsIgnoreCase(testCaseName)){
				switch (browserSelection) {
				case "firefox": 
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_FIREFOX, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				case "chrome":
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_CHROME, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				case "ie":
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_IE, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				case "safari":
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_SAFARI, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				case "opera":
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_OPERA, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				case "headless":
					if(xls.getCellData(Address_DataSheet_Constants.TESTCASES_WORKSHEET, Address_DataSheet_Constants.RUN_ON_HEADLESS, i).equalsIgnoreCase(Address_DataSheet_Constants.YES)){
						selectedBrowser = true;
					}
					break;
				default:
					System.out.println("The browser you have mentioned in xml file suite file is not updated in test data excel sheet");
					break;
				}
				//throw new SkipException("The Selenium script is not available to execute the test case "+testCaseName+ " mentioned in test case xlsx file. Hence skipped the test case execution");
			}
		}
		return selectedBrowser;
	}

	
	// return the test data from xls
	/**
	 * @param xls
	 * @param testCaseName
	 * @return
	 */
	public static Object[][] getData(Xls_ReadWrite xls, String testCaseName) {
		// if the sheet is not present
		if (!xls.isSheetExist(testCaseName)) {
			// throw new SkipException("The test data worksheet is not availble
			// for test case "+testCaseName+". HEnce skipped the test case
			// execution.");
			xls = null;
			return null;

		}

		int rows = xls.getRowCount(testCaseName);
		int cols = xls.getColumnCount(testCaseName);

		Object[][] data = new Object[rows - 1][cols];
		for (int rowNum = 2; rowNum <= rows; rowNum++) {
			for (int colNum = 0; colNum < cols; colNum++) {
				// System.out.print(xls.getCellData(testCaseName, colNum,
				// rowNum) + " -- ");
				data[rowNum - 2][colNum] = xls.getCellData(testCaseName, colNum, rowNum);
			}
		}
		return data;
	}

	// return the test data from xls
	/**
	 * @param xls
	 * @param testCaseName
	 * @param TDfields
	 * @return
	 */
	public static Object[][] getTestData(Xls_ReadWrite xls, String testCaseName, String[] TDfields,String TestData_WORKSHEET,String TESTCASES_TESTCASEID) {
		// if the sheet is not present
		if (!xls.isSheetExist(TestData_WORKSHEET)) {
			xls = null;
			return null;
//			throw new SkipException("The test data worksheet is not availble for test case "+testCaseName+". HEnce skipped the test caseexecution.");
			
		}

		int cols = TDfields.length;
		// System.out.println("number of input data are :"+cols);
		Object[][] data = new Object[1][cols];
		for (int i = 2; i <= xls.getRowCount(TestData_WORKSHEET); i++) {
			if (xls.getCellData(TestData_WORKSHEET, TESTCASES_TESTCASEID, i).equalsIgnoreCase(testCaseName)) {
				for (int colNum = 0; colNum < cols; colNum++) {
					if (xls.getCellData(TestData_WORKSHEET, TDfields[colNum], i).equals("@null")){
						System.out.println("The column name \'"+TDfields[colNum]+"\' is not available in test data excel sheet");
					}
					data[0][colNum] = xls.getCellData(TestData_WORKSHEET, TDfields[colNum], i);
				}break;
			}
		}
		return data;
	}	
	
	/**
	 * Author:- Sai Kiran Nataraja
	 * Get test data based on the test case name passed
	 * @param xls
	 * @param testCaseName
	 * @return
	
	 */
	public static  Object[][] getTestDataBasedOnTestCase(Xls_ReadWrite xls, String testCaseName) {
		try{
		// if the sheet is not present
		if (!xls.isSheetExist(Address_DataSheet_Constants.TestData_WORKSHEET)) {
			xls = null;
			throw new SkipException("The test data worksheet is not available for test case "+testCaseName+". Hence skipped the test case execution.");
			}
		}catch(Exception e){
			return null;	
		}
		
		int rows = xls.getRowCount(Address_DataSheet_Constants.TestData_WORKSHEET);
		
		for (int i = 2; i <= rows; i++) {
			if(xls.getCellData(Address_DataSheet_Constants.TestData_WORKSHEET, Address_DataSheet_Constants.TESTCASES_TESTCASEID, i).equalsIgnoreCase(testCaseName)){
				rows=i;
				break;
			}
		}
		
		int cols = xls.getColCountForParticularRow(Address_DataSheet_Constants.TestData_WORKSHEET,rows);
		int headerName=rows+1;
		int headerValue=rows+2;

//		System.out.println("Row number found: "+rows);
//		System.out.println("Col number found: "+cols);
		
		
		
		String[][] TDataFields = new String[1][cols];
		for(int colNumber=0;colNumber<cols;colNumber++){
			if(!xls.getCellData(Address_DataSheet_Constants.TestData_WORKSHEET, colNumber, headerName).equals("")){
//			System.out.println("Header Name: "+xls.getCellData(Address_DataSheet_Constants.TestData_WORKSHEET, colNumber,headerName ));
//			System.out.println("Header Value: "+xls.getCellData(Address_DataSheet_Constants.TestData_WORKSHEET, colNumber, headerValue));
			TDataFields[0][colNumber]=xls.getCellData(Address_DataSheet_Constants.TestData_WORKSHEET, colNumber, headerValue);
			}
		}
		return TDataFields;		
	}	
	
	
		
	
}
