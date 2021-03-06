package testUtilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.channels.OverlappingFileLockException;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import configurationFiles.Address_DataSheet_Constants;
import configurationFiles.CONSTANTS;


public class Xls_ReadWrite {
	public  String path;
	public  FileInputStream fis = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row   =null;
	public XSSFCell cell = null;
	public XSSFCell HistoryCell = null;

	/**
	 * @param path
	 */
	public Xls_ReadWrite(String path) {
		//		System.gc();
		this.path=path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {

			e.printStackTrace();
		} 

	}
	
	/**
	 * @author saikiran.nataraja
	 * @param sheetName
	 * @returns 0 - if index is -1, else returns the last row number
	 */
	public int getRowCount(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return 0;
		else{
			sheet = workbook.getSheetAt(index);
			int number=sheet.getLastRowNum()+1;
			return number;
		}
	}

	/**
	 * Author:- Sai Kiran Nataraja
	 * Get Column Count for a particular row in a worksheet 
	 * @param sheetName
	 * @param rowNumber
	 * @return
	 */
	/**
	 * @param sheetName
	 * @param rowNumber
	 * @return
	 */
	public int getColCountForParticularRow(String sheetName,int rowNumber){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1)
			return 0;
		else{
			sheet = workbook.getSheetAt(index);
			int noOfColumns = sheet.getRow(rowNumber).getPhysicalNumberOfCells();
			//getLastCellNum();
			return noOfColumns;
		}

	}

	// returns the data from a cell
	/**
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @return
	 */
	public String getCellData(String sheetName,String colName,int rowNum){
		try{
			if(rowNum <=0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			int col_Num=-1;
			if(index==-1)
				return "";

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);
			for(int i=0;i<row.getLastCellNum();i++){
				//System.out.println(row.getCell(i).getStringCellValue().trim());
				if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num=i;
			}
			if(col_Num==-1)
				return "@null";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum-1);
			if(row==null)
				return "";
			cell = row.getCell(col_Num);

			if(cell==null)
				return "";
			//System.out.println(cell.getCellType());
			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){

				String cellText  = String.valueOf(cell.getNumericCellValue());
				if (DateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal =Calendar.getInstance();
					cal.setTime(DateUtil.getJavaDate(d));
					cellText =
							(String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" +
							cal.get(Calendar.MONTH)+1 + "/" + 
							cellText;

					//System.out.println(cellText);

				}


				return cellText;
			}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return ""; 
			else 
				return String.valueOf(cell.getBooleanCellValue());

		}
		catch(Exception e){

			e.printStackTrace();
			return "row "+rowNum+" or column "+colName +" does not exist in xls";
		}
	}

	// returns the data from a cell
	/**
	 * @param sheetName
	 * @param colNum
	 * @param rowNum
	 * @return
	 */
	public String getCellData(String sheetName,int colNum,int rowNum){
		try{
			if(rowNum <=0)
				return "";
			int index = workbook.getSheetIndex(sheetName);
			if(index==-1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum-1);
			if(row==null)
				return "";
			cell = row.getCell(colNum);
			if(cell==null)
				return "";

			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
				//System.out.println("in xls read. befor :"+cell.getNumericCellValue());
				String cellText  = String.valueOf(cell.getNumericCellValue());
				//System.out.println("in xls read. after :"+cellText);
				if (DateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal =Calendar.getInstance();
					cal.setTime(DateUtil.getJavaDate(d));
					cellText =
							(String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH)+1 + "/" +
							cal.get(Calendar.DAY_OF_MONTH) + "/" +
							cellText;

					// System.out.println(cellText);
				}
				return cellText;
			}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else 
				return String.valueOf(cell.getBooleanCellValue());
		}
		catch(Exception e){

			e.printStackTrace();
			return "row "+rowNum+" or column "+colNum +" does not exist  in xls";
		}
	}	


	/*// returns true if data is set successfully else false
	 *//**
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @param data
	 * @return
	 *//*
	public boolean setCellData(String sheetName, String colName, int rowNum, String data) {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			if (rowNum <= 0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				System.out.println(row.getCell(i).getStringCellValue().trim());
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;

			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			// CellStyle cs = workbook.createCellStyle();
			// cs.setWrapText(true);
			// cell.setCellStyle(cs);
			cell.setCellValue(data);
			FileOutputStream fileOut = new FileOutputStream(path);
			workbook.write(fileOut);

			fileOut.close();

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}*/


	/**
	 * Function to find the row number of the specific string available in the excel
	 * @author saikiran.nataraja
	 * @param cellContent
	 * @returns Row number of the String present, Otherwise returns 0
	 */
	public int findRowNumber(String cellContent) {
		sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
	    for (Row row : sheet) {
	        for (Cell cell : row) {
	            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(cellContent)) {
	                    return row.getRowNum();  
	                }
	            }
	        }
	    }               
	    return 0;
	}
	
	/**
	 * Function to set cell data in the spreadsheet
	 * @author saikiran.nataraja 
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @param data
	 * @returns true if data is set successfully else false
	 */
	public boolean setCellData(String sheetName,String colName,int rowNum, String data){
		try{
			fis = new FileInputStream(path); 
			workbook = new XSSFWorkbook(fis);
			XSSFCellStyle fail = workbook.createCellStyle();
			XSSFCellStyle pass = workbook.createCellStyle();
			XSSFCellStyle skip = workbook.createCellStyle();
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

			skip.setFillForegroundColor(HSSFColor.GOLD.index);
			skip.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
			// Set the border style for the workbook
			skip.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			skip.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			skip.setBorderRight(XSSFCellStyle.BORDER_THIN);
			skip.setBorderTop(XSSFCellStyle.BORDER_THIN);
			skip.setAlignment(XSSFCellStyle.ALIGN_LEFT);

			if(rowNum<=0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum=-1;
			if(index==-1)
				return false;

			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(0);

			DataFormatter formatter = new DataFormatter();
			for (Iterator<?> iterator = sheet.rowIterator(); iterator.hasNext();) {
				XSSFRow row = (XSSFRow) iterator.next();
				for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
					XSSFCell cell = row.getCell(i);
					if(formatter .formatCellValue(cell).equals(colName.trim())){
						colNum=i;
						//				    		System.out.println(formatter .formatCellValue(cell));
						//				    		System.out.println(colNum);
						break;
					}
				}
				if(colNum>0) {
					break;
				}
			}			

			if(colNum==-1)
				return false;

			sheet.autoSizeColumn(colNum); 
			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);

			cell = row.getCell(colNum);	
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			XSSFCellStyle TableContentsOnLeft = workbook.createCellStyle();
			// Set the Font details for the entire sheet
			XSSFFont defaultFont;
			defaultFont = workbook.createFont();
			defaultFont.setFontHeightInPoints((short) 11);
			defaultFont.setFontName("Calibri");
			defaultFont.setColor(IndexedColors.BLACK.getIndex());
			defaultFont.setBold(true);
			defaultFont.setItalic(false);

			// create style for cells in table contents
			TableContentsOnLeft.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
			TableContentsOnLeft.setFillPattern(XSSFCellStyle.NO_FILL);
			//				TableContentsOnLeft.setWrapText(false);
			// Set the border style for the workbook
			TableContentsOnLeft.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderRight(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setBorderTop(XSSFCellStyle.BORDER_THIN);
			TableContentsOnLeft.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			switch (colName) {
			case "Execution Status":
				switch (data) {
				case "PASSED":
					cell.setCellStyle(pass);
					break;

				case "FAILED":
					cell.setCellStyle(fail);
					break;
				case "NOT EXECUTED":
					cell.setCellStyle(skip);
					break;
				}
				break;

			default:
				cell.setCellStyle(TableContentsOnLeft);
			}

			cell.setCellValue(data);
			FileOutputStream fileOut = new FileOutputStream(path);
			workbook.write(fileOut);

			fileOut.close();	

		}
		catch(Exception e){
			e.printStackTrace();
			return false;
		}
		return true;
	}


	/**
	 * Function to set the data for a value from a test case
	 * @author saikiran.nataraja
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @param data
	 * @return
	 */
	public boolean setCellDataFromTestCase(String sheetName,String colName,String TestcaseName, String data){
		try{
			fis = new FileInputStream(path); 
			workbook = new XSSFWorkbook(fis);
			int rowNum = 0;
			for(int i=2; i<= CONSTANTS.TCxls.getRowCount(sheetName) ; i++){
				if(CONSTANTS.TCxls.getCellData(sheetName, Address_DataSheet_Constants.TESTCASES_TESTCASEID, i).equalsIgnoreCase(TestcaseName)){
					rowNum=i+1; //Pass the Header Row Number for that test case so that it can get the required header
					break;
				}
			}

			if(rowNum<=0)
				return false;

			int index = workbook.getSheetIndex(sheetName);
			int colNum=-1;
			if(index==-1)
				return false;
			sheet = workbook.getSheetAt(index);
			row=sheet.getRow(rowNum-1); //Get the Header Row Name as per test case
			//			System.out.println("Column Count is:"+getColCountForParticularRow(sheetName,rowNum));
			for(int i=0;i<getColCountForParticularRow(sheetName,rowNum);i++){
				//				System.out.println(row.getCell(i).getStringCellValue().trim());
				if(!(row.getCell(i).getStringCellValue().isEmpty())){
					if(row.getCell(i).getStringCellValue().trim().equals(colName)){
						colNum=i;
						break;
					}
				}
			}
			if(colNum==-1)
				return false;

			sheet.autoSizeColumn(colNum); 
			row = sheet.getRow(rowNum);//Get the Value of Header Row passed as per test case
			if (row == null)
				row = sheet.createRow(rowNum-1);

			cell = row.getCell(colNum);	
			if (cell == null)
				cell = row.createCell(colNum);

			// cell style
			CellStyle cs = workbook.createCellStyle();
			cs.setWrapText(true);
			cs.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(cs);
			cell.setCellValue(data);

			/*//History Management - To know the previously used test data
			CellStyle PrevCs = workbook.createCellStyle();
			PrevCs.setWrapText(true);
			PrevCs.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			PrevCs.setFillPattern(CellStyle.SOLID_FOREGROUND);
			rowNum=rowNum+1;
			row=sheet.getRow(rowNum);
			HistoryCell = row.getCell(colNum);	
			if(HistoryCell==null)
				HistoryCell=row.createCell(colNum);
			HistoryCell.setCellStyle(PrevCs);
			HistoryCell.setCellValue("Previously Used "+colName+"is: "+data);*/
			FileOutputStream fileOut = new FileOutputStream(path);
			workbook.write(fileOut);			
			fileOut.close();
			workbook.close();
			fis.close();
		}catch (OverlappingFileLockException e) {
			// File is open by someone else
		}catch(Exception e){
			//			e.printStackTrace();
			return false;
		}
		return true;
	}


	// find whether sheets exists	
	/**
	 * @param sheetName
	 * @return
	 */
	public boolean isSheetExist(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		if(index==-1){
			index=workbook.getSheetIndex(sheetName.toUpperCase());
			if(index==-1)
				return false;
			else
				return true;
		}
		else
			return true;
	}

	// returns number of columns in a sheet	
	/**
	 * @param sheetName
	 * @return
	 */
	public int getColumnCount(String sheetName){
		// check if sheet exists
		if(!isSheetExist(sheetName))
			return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);

		if(row==null)
			return -1;

		return row.getLastCellNum();
		//		return row.getPhysicalNumberOfCells();


	}
}
