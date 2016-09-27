/*'*************************************************************************************************************************************************
' Script Name			: ExcelReportGenerator
' Description			: Used to convert TestNG report to ExcelReportGenerator
' How to Use			:
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Author                    Version          Creation Date         Reviewer Name         Reviewed Date           Comments 
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Sai Kiran Nataraja         v1.0             21-Sep-2016
'*************************************************************************************************************************************************
 */

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import fi.tapiola.stsovellukset01.config.CONSTANTS;

public class XLReportGenerator {

	private XLReportGenerator() {
		// To hide the public class we are making a constructor
	}

	public void generateExcelReport(String destFileName)
			throws ParserConfigurationException, SAXException, IOException {
		String path = XLReportGenerator.class.getClassLoader().getResource("./").getPath();
		path = path.replace("bin", "src");

		File xmlResultFile = new File(path + "../test-output/testng-results.xml");
		// System.out.println(xmlResultFile.isFile()); //--> To check if the
		// File exists or not

		// To Parse the XML Results file using the Document Builder factory of
		// javax jars
		DocumentBuilderFactory fact = DocumentBuilderFactory.newInstance();
		DocumentBuilder build = fact.newDocumentBuilder();
		Document doc = build.parse(xmlResultFile);
		doc.getDocumentElement().normalize();

		// Create an excel workbook
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFCellStyle fail = book.createCellStyle();
		XSSFCellStyle pass = book.createCellStyle();
		XSSFCellStyle TableHeader = book.createCellStyle();
		XSSFCellStyle TableContents = book.createCellStyle();

		XSSFFont defaultFont = book.createFont();
		defaultFont.setFontHeightInPoints((short) 11);
		defaultFont.setFontName("Calibri");
		defaultFont.setColor(IndexedColors.BLACK.getIndex());
		defaultFont.setBold(true);
		defaultFont.setItalic(false);

		// create style for cells in header row
		TableHeader.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
		TableHeader.setFont(defaultFont);
		TableHeader.setFillPattern(XSSFCellStyle.DIAMONDS);
		// Set the border style for the workbook
		TableHeader.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		TableHeader.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		TableHeader.setBorderRight(XSSFCellStyle.BORDER_THIN);
		TableHeader.setBorderTop(XSSFCellStyle.BORDER_THIN);
		TableHeader.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		// create style for cells in table contents
		TableContents.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
		TableContents.setFillPattern(XSSFCellStyle.NO_FILL);
		TableContents.setWrapText(true);
		// Set the border style for the workbook
		TableContents.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		TableContents.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		TableContents.setBorderRight(XSSFCellStyle.BORDER_THIN);
		TableContents.setBorderTop(XSSFCellStyle.BORDER_THIN);
		TableContents.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		fail.setFillForegroundColor(HSSFColor.RED.index);
		fail.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		// Set the border style for the workbook
		fail.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		fail.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		fail.setBorderRight(XSSFCellStyle.BORDER_THIN);
		fail.setBorderTop(XSSFCellStyle.BORDER_THIN);
		fail.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		pass.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
		pass.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		// Set the border style for the workbook
		pass.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		pass.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		pass.setBorderRight(XSSFCellStyle.BORDER_THIN);
		pass.setBorderTop(XSSFCellStyle.BORDER_THIN);
		pass.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		NodeList Suite_List = doc.getElementsByTagName("suite");
		for (int suits = 0; suits < Suite_List.getLength(); suits++) {
			int rowNum = 0;
			Node suits_Node = Suite_List.item(0);
			String suit_Name = ((Element) suits_Node).getAttribute("name");
			XSSFSheet sheet = book.createSheet(suit_Name);
			XSSFRow row = sheet.createRow(rowNum++);
			XSSFCell cel_name = row.createCell(0);
			XSSFCell cel_status = row.createCell(1);
			XSSFCell cel_exp = row.createCell(2);
			// Set the Header names for the sheet
			cel_name.setCellValue("Name of the Test case");
			cel_status.setCellValue("Status");
			cel_exp.setCellValue("Comments");
			// set style for the table header
			cel_name.setCellStyle(TableHeader);
			cel_status.setCellStyle(TableHeader);
			cel_exp.setCellStyle(TableHeader);
			NodeList Tests_List = doc.getElementsByTagName("test");
			for (int i = 0; i < Tests_List.getLength(); i++) {
				Node tests_Node = Tests_List.item(i);
				String Test_Name = ((Element) tests_Node).getAttribute("name");
				NodeList Class_List = ((Element) tests_Node).getElementsByTagName("class");
				for (int j = 0; j < Class_List.getLength(); j++) {
					Node classes_Node = Class_List.item(j);
					String Class_Name = ((Element) classes_Node).getAttribute("name");
					Class_Name = Class_Name.substring(Class_Name.lastIndexOf('.') + 1);
					NodeList TestMethods_List = ((Element) classes_Node).getElementsByTagName("test-method");
					for (int k = 0; k < TestMethods_List.getLength(); k++) {
						Node testMethod_Node = TestMethods_List.item(k);
						String testMethod_Name = ((Element) testMethod_Node).getAttribute("name");
						String testMethod_Status = ((Element) testMethod_Node).getAttribute("status");
						// Capture only the test case number ignoring the
						// beforetest,after test and after class methods
						if (testMethod_Name.contains("TC") == true) {
							row = sheet.createRow(rowNum++);

							// Create a column for the test case name
							cel_name = row.createCell(0);
							cel_name.setCellValue(Test_Name.trim());

							// Create column for the test case status
							cel_status = row.createCell(1);
							cel_status.setCellValue(testMethod_Status.trim());

							// Create column for the test case comments (if any)
							cel_exp = row.createCell(2);

							if ("fail".equalsIgnoreCase(testMethod_Status)) {
								NodeList exp_List = ((Element) testMethod_Node).getElementsByTagName("exception");
								Node exp_Node = exp_List.item(i);
								String exp_msg = ((Element) exp_Node).getAttribute("class");
								cel_status.setCellStyle(fail);
								cel_exp.setCellValue(exp_msg.trim());
							} else {
								cel_status.setCellStyle(pass);
							}
							// set style for the table contents
							cel_name.setCellStyle(TableContents);
							cel_exp.setCellStyle(TableContents);

							// Auto size the column widths based on the names
							sheet.autoSizeColumn(0);
							sheet.autoSizeColumn(1);
							sheet.autoSizeColumn(2);
						}
					}
				}
			}
		}
		FileOutputStream writeXLOutput = new FileOutputStream(CONSTANTS.ScreenshotsPath + destFileName);
		book.write(writeXLOutput);
		writeXLOutput.close();
		book.close();
		System.out.println("Report is Generated");
	}

	public static void main(String[] args) {
		try {
			// new XLReportGenerator().generateExcelReport(CONSTANTS.Dateformat
			// + "ExcelAutomationReport.xlsx");
			new XLReportGenerator().generateExcelReport("ExcelAutomationReport.xlsx");
		} catch (ParserConfigurationException | SAXException | IOException e) {
			e.printStackTrace();
		}
	}
}
