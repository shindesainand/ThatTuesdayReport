package com.consolidate.auto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.tracks.names.Constants;

public class MetricsDataCollector
{
	private final String CHARTS_FILE_NAME = "C:\\Users\\sainand\\Music\\WSRTempReport.xlsx";
	FileInputStream excelFile;
	Workbook workbook;
	Sheet sheet;
	WSRDataParser wsr;
	int compBenMgmntCol = 1;
	int workforceMgmntCol = 2;
	int hrCoreCol = 7;
	int efsPayrollCol = 3;
	int efsCol = 4;
	int wssCol = 5;
	int workforceDataAnalyticsCol = 6;
	int employeeLearningCol = 8;
	int externalLearningCol = 9;
	
	HashMap<String, ArrayList<Integer>> weekCaseMetrics = new HashMap<String, ArrayList<Integer>>();
	HashMap<String, ArrayList<Integer>> caseAgeMetrics = new HashMap<String, ArrayList<Integer>>();
	HashMap<String, ArrayList<Integer>> quarterCaseMetrics = new HashMap<String, ArrayList<Integer>>();
	HashMap<String, ArrayList<Integer>> requestRestoreMetrics = new HashMap<String, ArrayList<Integer>>();
	HashMap<String, ArrayList<Integer>> pbiMetrics = new HashMap<String, ArrayList<Integer>>();
	HashMap<String, ArrayList<Double>> mttr70thPercentileMetrics = new HashMap<String, ArrayList<Double>>();
	
	MetricsDataCollector()
	{
		try
		{
			excelFile = new FileInputStream(new File(CHARTS_FILE_NAME));
			workbook = new XSSFWorkbook(excelFile);
			sheet = workbook.getSheetAt(0);
			wsr = new WSRDataParser();
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
	
	public static void main(String[] args)
	{
		MetricsDataCollector dataCollector = new MetricsDataCollector();
		
		dataCollector.weekCaseMetrics = dataCollector.makeWeekCaseMetricsGraph();
		dataCollector.caseAgeMetrics = dataCollector.makeCaseAgeGraph();
		dataCollector.quarterCaseMetrics = dataCollector.makeQuarterCaseMetricsGraph();
		dataCollector.requestRestoreMetrics = dataCollector.makeRequestRestoreGraph();
		dataCollector.pbiMetrics = dataCollector.makePBIMetricsGraph();
		dataCollector.mttr70thPercentileMetrics = dataCollector.make70thPercentileMTTRGraph();
		
		System.out.println(dataCollector.weekCaseMetrics);
		System.out.println(dataCollector.caseAgeMetrics);
		System.out.println(dataCollector.quarterCaseMetrics);
		System.out.println(dataCollector.requestRestoreMetrics);
		System.out.println(dataCollector.pbiMetrics);
		System.out.println(dataCollector.mttr70thPercentileMetrics);
	}

	private HashMap<String, ArrayList<Integer>> makePBIMetricsGraph()
	{
		int beginRow = 25;
		int endRow = 27;
		HashMap<String, ArrayList<Integer>> hashMap = new HashMap<String, ArrayList<Integer>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmntAndHRCore, wsr.getWorkforceMgmntHRCoreData(workforceMgmntCol, hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.efsAndEfsPayroll, wsr.getEFSData(efsPayrollCol, efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}

	private HashMap<String, ArrayList<Integer>> makeRequestRestoreGraph()
	{
		int beginRow = 21;
		int endRow = 22;
		HashMap<String, ArrayList<Integer>> hashMap = new HashMap<String, ArrayList<Integer>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmntAndHRCore, wsr.getWorkforceMgmntHRCoreData(workforceMgmntCol, hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.efsAndEfsPayroll, wsr.getEFSData(efsPayrollCol, efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}

	private HashMap<String, ArrayList<Integer>> makeQuarterCaseMetricsGraph()
	{
		int beginRow = 16;
		int endRow = 18;
		HashMap<String, ArrayList<Integer>> hashMap = new HashMap<String, ArrayList<Integer>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmntAndHRCore, wsr.getWorkforceMgmntHRCoreData(workforceMgmntCol, hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.efsAndEfsPayroll, wsr.getEFSData(efsPayrollCol, efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}

	private HashMap<String, ArrayList<Integer>> makeCaseAgeGraph()
	{
		int beginRow = 8;
		int endRow = 12;
		HashMap<String, ArrayList<Integer>> hashMap = new HashMap<String, ArrayList<Integer>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmntAndHRCore, wsr.getWorkforceMgmntHRCoreData(workforceMgmntCol, hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.efsAndEfsPayroll, wsr.getEFSData(efsPayrollCol, efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}

	private HashMap<String, ArrayList<Integer>> makeWeekCaseMetricsGraph()
	{
		int beginRow = 2;
		int endRow = 4;
		HashMap<String, ArrayList<Integer>> hashMap = new HashMap<String, ArrayList<Integer>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmntAndHRCore, wsr.getWorkforceMgmntHRCoreData(workforceMgmntCol, hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.efsAndEfsPayroll, wsr.getEFSData(efsPayrollCol, efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}
	
	private HashMap<String, ArrayList<Double>> make70thPercentileMTTRGraph()
	{
		int beginRow = 34;
		int endRow = 35;
		HashMap<String, ArrayList<Double>> hashMap = new HashMap<String, ArrayList<Double>>();
		
		hashMap.put(Constants.compBenMgmnt, wsr.getCompBenMgmntMTTRData(compBenMgmntCol, beginRow, endRow));
		hashMap.put(Constants.workforceMgmnt, wsr.getWorkforceMgmntMTTRData(workforceMgmntCol, beginRow, endRow));
		hashMap.put(Constants.efsPayroll, wsr.getEFSPayrollMTTRData(efsPayrollCol, beginRow, endRow));
		hashMap.put(Constants.efs, wsr.getEFSMTTRData(efsCol, beginRow, endRow));
		hashMap.put(Constants.wss, wsr.getWSSMTTRData(wssCol, beginRow, endRow));
		hashMap.put(Constants.workforceDataAnalytics, wsr.getWorkforceDataAnalyticsMTTRData(workforceDataAnalyticsCol, beginRow, endRow));
		hashMap.put(Constants.hrCore, wsr.getHRCoreMTTRData(hrCoreCol, beginRow, endRow));
		hashMap.put(Constants.employeeLearning, wsr.getEmployeeLearningMTTRData(employeeLearningCol, beginRow, endRow));
		hashMap.put(Constants.externalLearning, wsr.getExternalLearningMTTRData(externalLearningCol, beginRow, endRow));
		
		return hashMap;
	}
}
