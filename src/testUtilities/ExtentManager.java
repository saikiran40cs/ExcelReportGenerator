package testUtilities;

import java.io.File;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import com.relevantcodes.extentreports.DisplayOrder;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.NetworkMode;

import configurationFiles.CONSTANTS;

public class ExtentManager {
    public static ExtentReports extent;
    protected static final String filePath=System.getProperty("user.dir")+File.separator+"ExtentReports"+File.separator+"extent.html";
	private static boolean replaceExisting=true;
	protected static final DisplayOrder displayOrder= DisplayOrder.NEWEST_FIRST;

    public static ExtentReports getInstance(String EnvironmentToRunOn) {
        if (extent == null) {
        	extent = new ExtentReports(filePath, replaceExisting,displayOrder, NetworkMode.OFFLINE,new Locale("fi_FI")); //en-US
        	extent.loadConfig(new File(CONSTANTS.ExtentConfigPath));
//    		String CurrentDate=new SimpleDateFormat("dd.MM.yyyy , HH:mm:ss").format(new Date(((new Date()).getTime() + 86400000)));
    		Map<String, String> sysInfo = new HashMap<String, String>();
    		sysInfo.put("Selenium Version", "2.51");
    		sysInfo.put("Environment", EnvironmentToRunOn);
    		extent.addSystemInfo(sysInfo);    		
        }
        return extent;
    }
}