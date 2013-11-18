package com.t.fang.common;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class FileUtil {
	private static final Logger logger = LoggerFactory.getLogger(FileUtil.class);

	public static void writeString(String path, String buf) {
		BufferedWriter bw = null;
		File file = new File(path);
		if (file.getParentFile().exists() == false) {
			file.getParentFile().mkdirs();
		}
		try {
			// 根据文件路径创建缓冲输出流
			bw = new BufferedWriter(new FileWriter(file));
			// 将内容写入文件中
			bw.write(buf);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			// 关闭流
			if (bw != null) {
				try {
					bw.close();
				} catch (IOException e) {
					bw = null;
				}
			}
		}
	}
	public static List readExcelToList(String excelPath, int firstRowNum) throws InvalidFormatException, IOException {
		ArrayList<List> rowsList = new ArrayList<List>();
		File source = new File(excelPath);
		Workbook workbook = null;
		DataFormatter formatter = null;
		FormulaEvaluator evaluator = null;
		Sheet sheet = null;
		Row row = null;
		int lastRowNum = 0;
		FileInputStream fis = null;
		try {
			System.out.println("Opening workbook [" + source.getName() + "]");

			fis = new FileInputStream(source);

			// Open the workbook and then create the FormulaEvaluator and
			// DataFormatter instances that will be needed to, respectively,
			// force evaluation of forumlae found in cells and create a
			// formatted String encapsulating the cells contents.
			workbook = WorkbookFactory.create(fis);
			evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			formatter = new DataFormatter(true);
		} catch (InvalidFormatException e) {
			logger.error("Invalid Format!", e);
			throw new InvalidFormatException(e.getMessage());
		} catch (IOException e) {
			logger.error("IOException!", e);
			throw new InvalidFormatException(e.getMessage());
		} finally {
			if (fis != null) {
				fis.close();
			}
		}

		int numSheets = workbook.getNumberOfSheets();

		// and then iterate through them.
		for (int i = 0; i < numSheets; i++) {

			// Get a reference to a sheet and check to see if it contains
			// any rows.
			sheet = workbook.getSheetAt(i);
			if (sheet.getPhysicalNumberOfRows() > 0) {

				// Note down the index number of the bottom-most row and
				// then iterate through all of the rows on the sheet starting
				// from the very first row - number 1 - even if it is missing.
				// Recover a reference to the row and then call another method
				// which will strip the data from the cells and build lines
				// for inclusion in the resylting CSV file.
				lastRowNum = sheet.getLastRowNum();
				for (int j = 0; j <= lastRowNum; j++) {
					row = sheet.getRow(j);
					Cell cell = null;
					int lastCellNum = 0;
					ArrayList<String> rowList = new ArrayList<String>();

					// Check to ensure that a row was recovered from the sheet as it is
					// possible that one or more rows between other populated rows could be
					// missing - blank. If the row does contain cells then...
						if (row != null) {
	
							// Get the index for the right most cell on the row and then
							// step along the row from left to right recovering the contents
							// of each cell, converting that into a formatted String and
							// then storing the String into the csvLine ArrayList.
							lastCellNum = row.getLastCellNum();
							for (int n = 0; n <= lastCellNum; n++) {
								cell = row.getCell(n);
								if (cell == null) {
									rowList.add("");
								} else {
									if (cell.getCellType() != Cell.CELL_TYPE_FORMULA) {
										rowList.add(formatter.formatCellValue(cell));
									} else {
										rowList.add(formatter.formatCellValue(cell, evaluator));
									}
								}
							}
							rowsList.add(rowList);
						}
				  }
			}
		}
		
		for(int i =0;i<rowsList.size();i++)
		{
			ArrayList<String> rowList = (ArrayList<String>) rowsList.get(i);
			for(int j=0;j<rowList.size();j++)
			{
				System.out.print(rowList.get(j)+',');
			}
			System.out.println();
		}
		return rowsList;
	}
	public static void saveListToExcel(ArrayList<List> rowsList) throws IOException{
		
		saveListToExcel(rowsList,null,null);
        
	}
	public static void saveListToExcel(ArrayList<List> rowsList,String fileName,String fileExtention) throws IOException{
		File newFile = new File("./File/testexcel.xls");
		//Workbook workbook = new  XSSFWorkbook();
		Workbook workbook = new  HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row row=null;
		ArrayList<String> rowList =null;
		for(int i =0;i<rowsList.size();i++)
		{   
		    rowList = (ArrayList<String>) rowsList.get(i);
		    row = sheet.createRow(i);
			Cell cell =null;
			for(int j=0;j<rowList.size();j++)
			{
			    cell = row.createCell(j);
			    cell.setCellType(1);
			    //cell.setCellStyle(style)
			    cell.setCellValue((String)rowList.get(j));
			}
		}
		 FileOutputStream fileOut =null;
		try {
			fileOut = new FileOutputStream(newFile);
			 workbook.write(fileOut);
			 fileOut.close();
		} catch (FileNotFoundException e) {
			logger.error("File Not Found!", e);
			throw new FileNotFoundException(e.getMessage());
		} catch (IOException e) {
			logger.error("IOExceptiont!", e);
			throw new IOException(e.getMessage());
		}finally{
			if(fileOut !=null){
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					fileOut=null;
				}
			}
		}
	    
	}

}
