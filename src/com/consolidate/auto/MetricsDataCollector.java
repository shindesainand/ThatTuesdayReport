package com.consolidate.auto;

import java.util.ArrayList;
import java.util.HashMap;

import com.tracks.names.Constants;

public class MetricsDataCollector
{
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
	
	public HashMap<String, ArrayList<Integer>> weekCaseMetrics = new HashMap<String, ArrayList<Integer>>();
	public HashMap<String, ArrayList<Integer>> caseAgeMetrics = new HashMap<String, ArrayList<Integer>>();
	public HashMap<String, ArrayList<Integer>> quarterCaseMetrics = new HashMap<String, ArrayList<Integer>>();
	public HashMap<String, ArrayList<Integer>> requestRestoreMetrics = new HashMap<String, ArrayList<Integer>>();
	public HashMap<String, ArrayList<Integer>> pbiMetrics = new HashMap<String, ArrayList<Integer>>();
	public HashMap<String, ArrayList<Double>> mttr70thPercentileMetrics = new HashMap<String, ArrayList<Double>>();
	
	public MetricsDataCollector()
	{
		wsr = new WSRDataParser();
		
		this.weekCaseMetrics = this.getWeekCaseMetrics();
		this.caseAgeMetrics = this.getCaseAgeMetrics();
		this.quarterCaseMetrics = this.getQuarterCaseMetrics();
		this.requestRestoreMetrics = this.getRequestRestoreMetrics();
		this.pbiMetrics = this.getPBIMetrics();
		this.mttr70thPercentileMetrics = this.get70thPercentileMTTRMetrics();
	}
	
	private HashMap<String, ArrayList<Integer>> getPBIMetrics()
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

	private HashMap<String, ArrayList<Integer>> getRequestRestoreMetrics()
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

	private HashMap<String, ArrayList<Integer>> getQuarterCaseMetrics()
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

	private HashMap<String, ArrayList<Integer>> getCaseAgeMetrics()
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

	private HashMap<String, ArrayList<Integer>> getWeekCaseMetrics()
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
	
	private HashMap<String, ArrayList<Double>> get70thPercentileMTTRMetrics()
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
