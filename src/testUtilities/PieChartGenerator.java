package testUtilities;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;


public class PieChartGenerator {

	private static void usage() throws IOException{
		String ReportPath = System.getProperty("user.dir")+File.separator+"ExtentReports"+File.separator+ "ExcelAutomationReport.xlsx";
		String ReportPath1 = System.getProperty("user.dir")+File.separator+"ExtentReports"+File.separator+ "ExcelAutomationReport1.xlsx";

		BufferedReader modelReader = new BufferedReader(new FileReader(ReportPath));
		XMLSlideShow pptx = null;
		try {
			String chartTitle = modelReader.readLine();  // first line is chart title

			pptx = new XMLSlideShow(new FileInputStream(ReportPath1));
			XSLFSlide slide = pptx.getSlides().get(0);

			// find chart in the slide
			XSLFChart chart = null;
			for(POIXMLDocumentPart part : slide.getRelations()){
				if(part instanceof XSLFChart){
					chart = (XSLFChart) part;
					break;
				}
			}

			if(chart == null) throw new IllegalStateException("chart not found in the template");

			// embedded Excel workbook that holds the chart data
			POIXMLDocumentPart xlsPart = chart.getRelations().get(0);
			XSSFWorkbook wb = new XSSFWorkbook();
			try {
				XSSFSheet sheet = wb.createSheet();

				CTChart ctChart = chart.getCTChart();
				CTPlotArea plotArea = ctChart.getPlotArea();

				CTPieChart pieChart = plotArea.getPieChartArray(0);
				//Pie Chart Series
				CTPieSer ser = pieChart.getSerArray(0);

				// Series Text
				CTSerTx tx = ser.getTx();
				tx.getStrRef().getStrCache().getPtArray(0).setV(chartTitle);
				sheet.createRow(0).createCell(1).setCellValue(chartTitle);
				String titleRef = new CellReference(sheet.getSheetName(), 0, 1, true, true).formatAsString();
				tx.getStrRef().setF(titleRef);

				// Category Axis Data
				CTAxDataSource cat = ser.getCat();
				CTStrData strData = cat.getStrRef().getStrCache();

				// Values
				CTNumDataSource val = ser.getVal();
				CTNumData numData = val.getNumRef().getNumCache();

				strData.setPtArray(null);  // unset old axis text
				numData.setPtArray(null);  // unset old values

				// set model
				int idx = 0;
				int rownum = 1;
				String ln;
				while((ln = modelReader.readLine()) != null){
					String[] vals = ln.split("\\s+");
					CTNumVal numVal = numData.addNewPt();
					numVal.setIdx(idx);
					numVal.setV(vals[1]);

					CTStrVal sVal = strData.addNewPt();
					sVal.setIdx(idx);
					sVal.setV(vals[0]);

					idx++;
					XSSFRow row = sheet.createRow(rownum++);
					row.createCell(0).setCellValue(vals[0]);
					row.createCell(1).setCellValue(Double.valueOf(vals[1]));
				}
				numData.getPtCount().setVal(idx);
				strData.getPtCount().setVal(idx);

				String numDataRange = new CellRangeAddress(1, rownum-1, 1, 1).formatAsString(sheet.getSheetName(), true);
				val.getNumRef().setF(numDataRange);
				String axisDataRange = new CellRangeAddress(1, rownum-1, 0, 0).formatAsString(sheet.getSheetName(), true);
				cat.getStrRef().setF(axisDataRange);

				// updated the embedded workbook with the data
				OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();
				try {
					wb.write(xlsOut);
				} finally {
					xlsOut.close();
				}

				// save the result
				OutputStream out = new FileOutputStream("pie-chart-demo-output.pptx");
				try {
					pptx.write(out);
				} finally {
					out.close();
				}
			} finally {
				wb.close();
			}
		} finally {
			if (pptx != null) pptx.close();
			modelReader.close();
		}

	}

	public static void main(String[] args) throws Exception {    	
		usage();
	}


}
