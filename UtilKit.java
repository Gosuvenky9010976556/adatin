package utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UtilKit {
	
	static String resourceFolder="src\\test\\resources";
	
	public static Object[][] getData(String sheetName,String testcase) throws IOException
	{
		
		FileInputStream fis1=new FileInputStream(resourceFolder+"\\excels\\MasterTestData.xlsx");
		
		XSSFWorkbook wb1=new XSSFWorkbook(fis1);
		
		XSSFSheet ws1=wb1.getSheet(sheetName);
		
		ArrayList<Row> testCaseRows=new ArrayList<Row>();
		
		testCaseRows=findRows(ws1,testcase);
		
		Object[][] obj=new Object[testCaseRows.size()-1][testCaseRows.get(0).getLastCellNum()-1];
		
		//Object[][] obj=Object[4][3];
		//obj[0][0]="reyaz0310";
		//obj[0][1]
		//obj[0][2]
		
		for(int i=0;i<testCaseRows.size()-1;i++)
		{
			
			for(int j=0;j<testCaseRows.get(0).getLastCellNum()-1;j++)
			{
				obj[i][j]=testCaseRows.get(i+1).getCell(j+1).getStringCellValue();
				
				System.out.print(obj[i][j]+"\t");
			}
			
			System.out.println();
		}
		
				
		return obj;
		
	}
	
	public static ArrayList<Row> findRows(Sheet ws,String testCase)
	{
		
		ArrayList<Row> rows=new ArrayList<Row>();
		
		ArrayList<Row> testCaserows=new ArrayList<Row>();
		
		System.out.println(ws.getLastRowNum());
		
		//To collect all rows having data only
		
		for(int i=0;i<=ws.getLastRowNum();i++)
		{
			
			if(!((ws.getRow(i))==null))
			{
				rows.add(ws.getRow(i));
			}
			
		}
		
		
		System.out.println(rows.size());
		
		
		//To collect rows for Test Case
		
		for(int i=0;i<rows.size();i++)
		{
			
			if(rows.get(i).getCell(0).getStringCellValue().equalsIgnoreCase(testCase))
			{
				testCaserows.add(rows.get(i));
			}
		}
		
		
		
		return testCaserows;
		
		
		
	}

}
