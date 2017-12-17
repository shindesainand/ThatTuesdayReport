package com.graphs.auto;

import java.util.ArrayList;
import java.util.HashMap;

import com.consolidate.auto.MetricsDataCollector;

public class GraphsMaker
{
	GraphSheetWriter graphWriter;
	
	GraphsMaker()
	{
		this.graphWriter = new GraphSheetWriter();
	}
	
	public static void main(String[] args)
	{
		MetricsDataCollector dataCollector = new MetricsDataCollector();
		
		GraphsMaker graphMaker = new GraphsMaker();
		
		System.out.println(dataCollector.weekCaseMetrics);
		graphMaker.makeWeekCaseMetricsGraph(dataCollector.weekCaseMetrics);

		System.out.println(dataCollector.caseAgeMetrics);
		graphMaker.makeCaseAgeMetricsGraph(dataCollector.caseAgeMetrics);
		
		System.out.println(dataCollector.quarterCaseMetrics);
		graphMaker.makeQuarterCaseMetricsGraph(dataCollector.quarterCaseMetrics);

		/*System.out.println(dataCollector.requestRestoreMetrics);
		graphMaker.makeReqRestoreMetricsGraph(dataCollector.requestRestoreMetrics);*/
		
		System.out.println(dataCollector.pbiMetrics);
		graphMaker.makePBIMetricsGraph(dataCollector.pbiMetrics);
		
		System.out.println(dataCollector.mttr70thPercentileMetrics);
		graphMaker.makeMTTR70thPercentileMetricsGraph(dataCollector.mttr70thPercentileMetrics);
	}

	private void makeWeekCaseMetricsGraph(HashMap<String, ArrayList<Integer>> weekCaseMetrics)
	{
		graphWriter.writeWeekCaseMetricsGraph(weekCaseMetrics);
	}
	
	private void makeCaseAgeMetricsGraph(HashMap<String, ArrayList<Integer>> caseAgeMetrics)
	{
		graphWriter.writeCaseAgeMetricsGraph(caseAgeMetrics);
	}
	
	private void makeQuarterCaseMetricsGraph(HashMap<String, ArrayList<Integer>> quarterCaseMetrics)
	{
		graphWriter.writeQuarterCaseMetricsGraph(quarterCaseMetrics);
	}
	
	private void makeReqRestoreMetricsGraph(HashMap<String, ArrayList<Integer>> requestRestoreMetrics)
	{
		graphWriter.writeReqRestoreMetricsGraph(requestRestoreMetrics);
	}
	
	private void makePBIMetricsGraph(HashMap<String, ArrayList<Integer>> pbiMetrics)
	{
		graphWriter.writePBIMetricsGraph(pbiMetrics);
	}
	
	private void makeMTTR70thPercentileMetricsGraph(HashMap<String, ArrayList<Double>> mttrMetrics)
	{
		graphWriter.writeMTTR70thPercentileMetricsGraph(mttrMetrics);
	}
}
