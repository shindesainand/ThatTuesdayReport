package com.bo.auto;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BaseDataParser {

	private final String BASE_FILE_NAME = "C:\\Users\\saishind\\Pictures\\Incident Details Overall_R8_final_L1_L2_combined.xlsx";
	FileInputStream excelFile;
	Workbook workbook;
	Sheet sheet;
	static String WSS[] = {"GSE-CVC-FIN-WPR", "GSE-CVC-FIN-SSBR", "GSE-L1-FM-WSS"};
	static String EFS[] = {"GSE-CVC-FIN-EFS-EMPSERV", "GSE-CVC-FIN-EFS-STOCK", "GSE-L1-CF-EFS"};
	static String yearWeekQuarter = "FY2017 Q4 WK07";
	static String yearQuarter = "FY2017 Q4";
	
	BaseDataParser()
	{
		try 
		{
			excelFile = new FileInputStream(new File(BASE_FILE_NAME));
			workbook = new XSSFWorkbook(excelFile);
			sheet = workbook.getSheetAt(7);
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
		BaseDataParser bdParser = new BaseDataParser();
		PBIParser pbiParser = new PBIParser();
		String tracks[][] = {WSS, EFS};
		
		for(int i = 0; i < tracks.length; i++)
		{
			System.out.println("METRICS FOR : "+tracks[i][i] +" and co-tracks");
			System.out.println("Week metrics: ");
			System.out.println(bdParser.getWeekCreated(tracks[i], yearWeekQuarter));
			System.out.println(bdParser.getWeekResolved(tracks[i], yearWeekQuarter));
			System.out.println(bdParser.getWeekBacklog(tracks[i]));
			
			System.out.println("Case age frequencies: ");
			for(int j = 0; j < 5; j++)
			{
				System.out.println(bdParser.getCaseAgeFreq(tracks[i], yearWeekQuarter)[i]);
			}
			System.out.println("Quarter metrics: ");
			System.out.println(bdParser.getQuarterCreated(tracks[i], yearQuarter));
			System.out.println(bdParser.getQuarterResolved(tracks[i], yearQuarter));
			System.out.println(bdParser.getQuarterBacklog(tracks[i]));
			
			System.out.println("Resolved case metrics: ");
			System.out.println(bdParser.getQuarterResolvedReq(tracks[i], yearQuarter));
			System.out.println(bdParser.getQuarterResolvedRestore(tracks[i], yearQuarter));
			
			System.out.println("PBI metrics: ");
			System.out.println(pbiParser.getCreated(tracks[i], yearQuarter));
			System.out.println(pbiParser.getResolved(tracks[i], yearQuarter));
			System.out.println(pbiParser.getBacklog(tracks[i])+"\n");
		}
	}

	private int getQuarterResolvedRestore(String[] assGroup, String yearQuarter)
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incResolvedWeek = curRow.getCell(27);
			Cell incStatus = curRow.getCell(21);
			Cell incServiceType = curRow.getCell(4);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incResolvedWeek != null)
							if(incResolvedWeek.getStringCellValue().contains(yearQuarter))
							{
								if(incStatus != null)
									if(incStatus.getStringCellValue().equals("Closed") || incStatus.getStringCellValue().equals("Resolved"))
									{
										if(incServiceType.getStringCellValue().equals("User Service Restoration"))
											count++;
									}
							}
					}
			}
		}
		return count;
	}
	
	private int getQuarterResolvedReq(String[] assGroup, String yearQuarter)
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incResolvedWeek = curRow.getCell(27);
			Cell incStatus = curRow.getCell(21);
			Cell incServiceType = curRow.getCell(4);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incResolvedWeek != null)
							if(incResolvedWeek.getStringCellValue().contains(yearQuarter))
							{
								if(incStatus != null)
									if(incStatus.getStringCellValue().equals("Closed") || incStatus.getStringCellValue().equals("Resolved"))
									{
										if(incServiceType.getStringCellValue().equals("User Service Request"))
											count++;
									}
							}
					}
			}
		}
		return count;
	}

	private int getQuarterBacklog(String[] assGroup)
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incStatus != null)
							if(incStatus.getStringCellValue().equals("Assigned") || incStatus.getStringCellValue().equals("In Progress") || incStatus.getStringCellValue().equals("Pending"))
								count++;
					}
			}
		}
		return count;
	}

	private int getQuarterResolved(String[] assGroup, String yearQuarter) 
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incResolvedWeek = curRow.getCell(27);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incResolvedWeek != null)
							if(incResolvedWeek.getStringCellValue().contains(yearQuarter))
							{
								if(incStatus != null)
									if(incStatus.getStringCellValue().equals("Closed") || incStatus.getStringCellValue().equals("Resolved"))
										count++;
							}
					}
			}
		}
		return count;
	}

	private int getQuarterCreated(String[] assGroup, String yearQuarter) 
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incSubmitWeek = curRow.getCell(25);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incSubmitWeek != null)
							if(incSubmitWeek.getStringCellValue().contains(yearQuarter))
							{
								if(incStatus != null)
									if(!incStatus.getStringCellValue().equals("Cancelled"))
										count++;
							}
					}
			}
		}
		return count;
	}

	private int[] getCaseAgeFreq(String[] assGroup, String yearWeekQuarter) 
	{
		int caseAgeFreq[] = new int[5];
		Iterator<Row> rowIterator = sheet.iterator();
		ArrayList<Double> ages = new ArrayList<Double>();
		DecimalFormat decimalFormat = new DecimalFormat("##.00");
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incSubmitWeek = curRow.getCell(25);
			Cell incStatus = curRow.getCell(21);
			Cell incAge = curRow.getCell(66);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incSubmitWeek != null)
							if(incSubmitWeek.getStringCellValue().equals(yearWeekQuarter))
							{
								if(incStatus != null)
									if(incStatus.getStringCellValue().equals("Assigned") || incStatus.getStringCellValue().equals("In Progress") || incStatus.getStringCellValue().equals("Pending"))
									{
										//System.out.println(deciFormat.format(incAge.getNumericCellValue()));
										ages.add(Double.parseDouble(decimalFormat.format(incAge.getNumericCellValue())));
									}
							}
					}
			}
		}

		for(int i = 0; i < ages.size(); i++)
		{
			double age = ages.get(i);
			
			if(age <= 2)
				caseAgeFreq[0]++;
			else if(age > 2 && age <= 5)
				caseAgeFreq[1]++;
			else if(age > 5 && age <= 13)
				caseAgeFreq[2]++;
			else if(age > 14 && age <= 29)
				caseAgeFreq[3]++;
			else if(age > 29 && age <= 59)
				caseAgeFreq[4]++;
		}
		
		return caseAgeFreq;
	}

	private int getWeekBacklog(String assGroup[]) 
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incStatus != null)
							if(incStatus.getStringCellValue().equals("Assigned") || incStatus.getStringCellValue().equals("In Progress") || incStatus.getStringCellValue().equals("Pending"))
								count++;
					}
			}
		}
		return count;
	}

	private int getWeekResolved(String assGroup[], String yearWeekQuarter)
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incResolvedWeek = curRow.getCell(27);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incResolvedWeek != null)
							if(incResolvedWeek.getStringCellValue().equals(yearWeekQuarter))
							{
								if(incStatus != null)
									if(incStatus.getStringCellValue().equals("Closed") || incStatus.getStringCellValue().equals("Resolved"))
										count++;
							}
					}
			}
		}
		return count;
	}

	private int getWeekCreated(String assGroup[], String yearWeekQuarter) 
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(56);
			Cell incSubmitWeek = curRow.getCell(25);
			Cell incStatus = curRow.getCell(21);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(incSubmitWeek != null)
							if(incSubmitWeek.getStringCellValue().equals(yearWeekQuarter))
							{
								if(incStatus != null)
									if(!incStatus.getStringCellValue().equals("Cancelled"))
										count++;
							}
					}
			}
		}
		return count;
	}
}
