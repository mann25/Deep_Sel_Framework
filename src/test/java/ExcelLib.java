import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class ExcelLib {
	

	String filepath = "C:\\Users\\admin1\\Desktop\\Test_sheet_1.xls" ;
    public String ExcelData(String SheetName , int rowNum,int cellNum) throws InvalidFormatException, IOException{
		
	
	FileInputStream fis = new FileInputStream(filepath);
	Workbook wb = WorkbookFactory.create(fis);
	Sheet sh = wb.getSheet(SheetName);
	
	//Write data
	sh.getRow(rowNum).getCell(cellNum).setCellValue("data to be entered");
	FileOutputStream fos = new FileOutputStream(filepath);
	wb.write(fos);
	
	String  data = sh.getRow(rowNum).getCell(cellNum).getStringCellValue();  //read data
	System.out.println(data);
    
    return data;	
	   
	   
	  
    	
    }
    
    }


