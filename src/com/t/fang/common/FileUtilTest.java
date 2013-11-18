package com.t.fang.common;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import junit.framework.TestCase;


public class FileUtilTest {
	
	public FileUtilTest(){
		
	}
@Test
public void testReadExcelToList(){
	String filepath1="./File/airports.xlsx";
	String filepath2="./File/caac.xls";
	String filepath3="D:/project_doc/caac_transform/Domestic Airline –Domestic flight.xls";
	String filepath4="D:/project_doc/MU_Transform/MU2013年夏秋季航班计划.xls";
	ArrayList<List> rowsList = new ArrayList<List>();
	try {
		rowsList=(ArrayList<List>) FileUtil.readExcelToList(filepath3, 0);
	} catch (InvalidFormatException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	
	try {
		FileUtil.saveListToExcel(rowsList);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
}
}
