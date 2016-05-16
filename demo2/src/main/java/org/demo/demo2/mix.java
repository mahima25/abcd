package org.demo.demo2;
/*to use sheets we need to import jxl.jar file
 this is for reading and writing*/

import org.openqa.selenium.*;
import org.openqa.selenium.firefox.*;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import java.util.concurrent.*;
import java.util.*;
import java.io.*;

import org.apache.commons.io.FileUtils;






import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class mix {
	@Test
		   public static void main() 
		      throws BiffException, IOException, WriteException
		      {
			   //WRITING
			   WritableWorkbook wwb;
			   wwb= Workbook.createWorkbook(new File(("mahima.xls")));
			   WritableSheet s= wwb.createSheet("My First Sheet", 0);
			   wwb.getSheet(0);
			   String[] array= {"USER NAME","mahima","nikita","shalini"};
			  addValue(array,s);
			  String[] array1={"MARKS","123","123","123"};
	          addNumber(array1,s);
			   wwb.write();
			   wwb.close();
			   
		      //READING
		      Workbook wb= Workbook.getWorkbook(new File("mahima.xls"));
		      Sheet sheet=wb.getSheet(0);
		     int rows= sheet.getRows();
		     int column= sheet.getColumns();
		     
		     for (int i=1;i<rows;i++)
		     {
		    	 for (int j=0;j<column;j++)
		    	 {
		    		 System.out.print("The contents Are "+sheet.getCell(j,i).getContents()+"\n");                                
		    	 }
		     }
		      
		      }
		   
		   public static void addValue(String[] array,WritableSheet s) throws RowsExceededException, WriteException
		   {
			   int j =0;
			   for(int i=1;i<5;i++)
			   {
				
				
				   Label label=new Label(0,i,array[j]);
				   j++;
				   s.addCell(label);
				  
				   }
		   }
		   public static void addNumber(String[] array1,WritableSheet s) throws RowsExceededException, WriteException
		   {
			   int j =0;
			   for(int i=1;i<5;i++)
			   {
				
				
				  Label label = new Label(1,i,array1[j]);
				   j++;
				   s.addCell(label);
				  
				   }
			   
		   }}
			 


	

