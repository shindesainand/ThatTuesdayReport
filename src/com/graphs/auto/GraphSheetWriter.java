package com.graphs.auto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tracks.names.Constants;

public class GraphSheetWriter
{
	
	 /* Change file paths
	 * Trailing whitespace after one of the track names in excel graphs file
	 * Change PBI metrics table position in graphs template file
	 * Rename Workforce Management in case age graph to "Workforce Mgmt"
	 * */
	
	private final String CHARTS_FILE_NAME = "C:\\Users\\saishind\\Pictures\\WSR Charts.xlsx";
	FileInputStream excelFile;
	Workbook workbook;
	Sheet sheet;

	GraphSheetWriter()
	{
		try 
		{
			this.excelFile = new FileInputStream(new File(CHARTS_FILE_NAME));
			this.workbook = new XSSFWorkbook(excelFile);
		}
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	void writeWeekCaseMetricsGraph(HashMap<String, ArrayList<Integer>> weekCaseMetrics)
	{
		int sheetNo = 0;
		int trackNameCol = 2;
		int startRow = 2;
		int endRow = 8;
		int startCol = 3;
		int endCol = 5;
		
		excelDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, weekCaseMetrics);
	}
	
	void writeCaseAgeMetricsGraph(HashMap<String, ArrayList<Integer>> caseAgeMetrics)
	{
		int sheetNo = 4;
		int trackNameCol = 14;
		int startRow = 5;
		int endRow = 11;
		int startCol = 23;
		int endCol = 27;
		
		excelDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, caseAgeMetrics);
	}
	
	void writeQuarterCaseMetricsGraph(HashMap<String, ArrayList<Integer>> quarterCaseMetrics)
	{
		int sheetNo = 1;
		int trackNameCol = 2;
		int startRow = 2;
		int endRow = 8;
		int startCol = 3;
		int endCol = 5;
		
		excelDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, quarterCaseMetrics);
	}
	
	void writeReqRestoreMetricsGraph(HashMap<String, ArrayList<Integer>> requestRestoreMetrics)
	{
		int sheetNo = 3;
		int trackNameCol = 2;
		int startRow = 2;
		int endRow = 8;
		int startCol = 3;
		int endCol = 4;
		
		excelDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, requestRestoreMetrics);
	}
	
	void writePBIMetricsGraph(HashMap<String, ArrayList<Integer>> pbiMetrics)
	{
		int sheetNo = 2;
		int trackNameCol = 2;
		int startRow = 2;
		int endRow = 8;
		int startCol = 3;
		int endCol = 5;
		
		excelDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, pbiMetrics);
	}
	
	void writeMTTR70thPercentileMetricsGraph(HashMap<String, ArrayList<Double>> mttrMetrics)
	{
		int sheetNo = 0;
		int trackNameCol = 10;
		int startRow = 2;
		int endRow = 8;
		int startCol = 11;
		int endCol = 12;
		
		excelMTTRDataWriterWorker(sheetNo, trackNameCol, startRow, endRow, startCol, endCol, mttrMetrics);
	}

	private void excelDataWriterWorker(int sheetNo, int trackNameCol, int beginRow, int endRow, int beginCol, int endCol, HashMap<String, ArrayList<Integer>> metrics)
	{
		sheet = workbook.getSheetAt(sheetNo);
		
		for(int i = beginRow; i <= endRow; i++)
		{
			Row curRow = sheet.getRow(i);
			for(int k = 0, j = beginCol; j <= endCol; j++, k++)
			{
				String trackName = sheet.getRow(i).getCell(trackNameCol).getStringCellValue();
				//System.out.println(trackName);
				int val = metrics.get(trackName).get(k);
				curRow.getCell(j).setCellValue(val);
			}
		}
		writeAndCloseFile();
	}
	
	private void excelMTTRDataWriterWorker(int sheetNo, int trackNameCol, int beginRow, int endRow, int beginCol, int endCol, HashMap<String, ArrayList<Double>> metrics)
	{
		sheet = workbook.getSheetAt(sheetNo);
		
		for(int i = beginRow; i <= endRow; i++)
		{
			Row curRow = sheet.getRow(i);
			for(int k = 0, j = beginCol; j <= endCol; j++, k++)
			{
				String trackName = sheet.getRow(i).getCell(trackNameCol).getStringCellValue();
				double val = metrics.get(trackName) == null ? 0 : metrics.get(trackName).get(k);
				if(trackName.equals(Constants.workforceMgmntAndHRCore))
				{
					val = 0;
				}
				System.out.println(trackName +" --- "+ val);
				curRow.getCell(j).setCellValue(val);
			}
		}
		writeAndCloseFile();
	}

	private void writeAndCloseFile()
	{
		try
		{
			FileOutputStream outFile = new FileOutputStream(new File(CHARTS_FILE_NAME));
			workbook.write(outFile);
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
}
