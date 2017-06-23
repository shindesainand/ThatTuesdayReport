package com.graphs.auto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WSRDataParser {
	
	private final String DATA_FILE_NAME = "C:\\Users\\saishind\\Pictures\\WSRTempReport.xlsx";
	FileInputStream excelFile;
	Workbook workbook;
	Sheet sheet;
	DecimalFormat decimalFormat = new DecimalFormat("##.00");

	WSRDataParser()
	{
		try
		{
			excelFile = new FileInputStream(new File(DATA_FILE_NAME));
			workbook = new XSSFWorkbook(excelFile);
			sheet = workbook.getSheetAt(0);
		}
		catch (FileNotFoundException e)
		{
			e.printStackTrace();
        }
		catch (IOException e) 
		{
            e.printStackTrace();
        }
	}
	
	ArrayList<Integer> getCompBenMgmntData(int c, int r1, int r2)
	{
		//(col, rowBegin, rowEnd)
		return excelDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Integer> getWorkforceMgmntHRCoreData(int c1, int c2, int r1, int r2)
	{
		ArrayList<Integer> workforceMgmnt = excelDataExtractorWorker(c1, r1, r2);
		ArrayList<Integer> hrCore = excelDataExtractorWorker(c2, r1, r2);
		return arrayListAdder(workforceMgmnt, hrCore);
	}
	
	ArrayList<Integer> getEFSData(int c1, int c2, int r1, int r2)
	{
		ArrayList<Integer> efsPayroll = excelDataExtractorWorker(c1, r1, r2);
		ArrayList<Integer> efs = excelDataExtractorWorker(c2, r1, r2);
		return arrayListAdder(efsPayroll, efs);
	}
	
	ArrayList<Integer> getWSSData(int c, int r1, int r2)
	{
		return excelDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Integer> getWorkforceDataAnalyticsData(int c, int r1, int r2)
	{
		return excelDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Integer> getEmployeeLearningData(int c, int r1, int r2)
	{
		return excelDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Integer> getExternalLearningData(int c, int r1, int r2)
	{
		return excelDataExtractorWorker(c, r1, r2);
	}
	
	private ArrayList<Integer> excelDataExtractorWorker(int col, int beginRow, int endRow)
	{
		ArrayList<Integer> array = new ArrayList<Integer>(); 
		for(int i = beginRow; i <= endRow; i++)
		{
			Row curRow = sheet.getRow(i);
			int cellVal = (int) curRow.getCell(col).getNumericCellValue();
			array.add(cellVal);
		}
		return array;
	}
	
	private ArrayList<Integer> arrayListAdder(ArrayList<Integer> list1, ArrayList<Integer> list2)
	{
		ArrayList<Integer> temp = new ArrayList<Integer>();
		for(int i = 0; i < list1.size(); i++)
		{
			temp.add(list1.get(i) + list2.get(i));
		}
		return temp;
	}
	/*
	 * 
	 * 
	 * 
	 * 
	 * */
	ArrayList<Double> getCompBenMgmntMTTRData(int c, int r1, int r2)
	{
		//(col, rowBegin, rowEnd)
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Double> getWorkforceMgmntMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Double> getHRCoreMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Double> getEFSPayrollMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Double> getEFSMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	ArrayList<Double> getWSSMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Double> getWorkforceDataAnalyticsMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Double> getEmployeeLearningMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}

	ArrayList<Double> getExternalLearningMTTRData(int c, int r1, int r2)
	{
		return excelMTTRDataExtractorWorker(c, r1, r2);
	}
	
	private ArrayList<Double> excelMTTRDataExtractorWorker(int col, int beginRow, int endRow)
	{
		ArrayList<Double> array = new ArrayList<Double>();
		for(int i = beginRow; i <= endRow; i++)
		{
			Row curRow = sheet.getRow(i);
			double cellVal = Double.parseDouble(decimalFormat.format(curRow.getCell(col).getNumericCellValue()));
			array.add(cellVal);
		}
		return array;
	}
}
