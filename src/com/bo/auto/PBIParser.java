package com.bo.auto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PBIParser {

	private final String BASE_FILE_NAME = "C:\\Users\\sainand\\Music\\Incident Details Overall_R8_final_L1_L2_combined.xlsx";
	FileInputStream excelFile;
	Workbook workbook;
	Sheet sheet;

	PBIParser()
	{
		try
		{
			excelFile = new FileInputStream(new File(BASE_FILE_NAME));
			workbook = new XSSFWorkbook(excelFile);
			sheet = workbook.getSheetAt(9);
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

	int getBacklog(String assGroup[])
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(42);
			Cell pbiStatus = curRow.getCell(41);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(pbiStatus != null)
							if(pbiStatus.getStringCellValue().equals("Draft") || pbiStatus.getStringCellValue().equals("Under Review") || pbiStatus.getStringCellValue().equals("Under Investigation") || pbiStatus.getStringCellValue().equals("Pending"))
								count++;
					}
			}
		}
		return count;
	}

	int getResolved(String assGroup[], String yearQuarter)
	{
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(42);
			Cell pbiCompleteQuarter = curRow.getCell(39);
			Cell pbiStatus = curRow.getCell(41);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(pbiCompleteQuarter != null)
							if(pbiCompleteQuarter.getStringCellValue().contains(yearQuarter))
							{
								if(pbiStatus != null)
									if(pbiStatus.getStringCellValue().equals("Completed") || pbiStatus.getStringCellValue().equals("Closed"))
										count++;
							}
					}
			}
		}
		return count;
	}

	int getCreated(String assGroup[], String yearQuarter) {
		int count = 0;
		Iterator<Row> rowIterator = sheet.iterator();
		
		while(rowIterator.hasNext())
		{
			Row curRow = rowIterator.next();
			Cell assignGrpName = curRow.getCell(42);
			Cell pbiSubmitQuarter = curRow.getCell(31);
			Cell pbiStatus = curRow.getCell(41);
			
			for(int i = 0; i < assGroup.length; i++)
			{
				if(assignGrpName != null)
					if(assignGrpName.getStringCellValue().equals(assGroup[i]))
					{
						if(pbiSubmitQuarter != null)
							if(pbiSubmitQuarter.getStringCellValue().equals(yearQuarter))
							{
								if(pbiStatus != null)
									if(!pbiStatus.getStringCellValue().equals("Cancelled"))
										count++;
							}
					}
			}
		}
		return count;
	}

}
