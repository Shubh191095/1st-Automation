package excelSheet;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;

public class autoExcelProg {

	
public static void main(String[] args) throws IOException {
//	String path = "C:\\excelData\\demo.xlsx";
	FileInputStream fs = new FileInputStream(path);

	XSSFWorkbook workbook= new XSSFWorkbook(fs);
	XSSFSheet sheet = workbook.getSheetAt(0);
	Row row = sheet.getRow(0);
	Workbook wb = new Workbook(fs);
	Sheet sheet1 = wb.getSheetAt(0);
	int lastRow = sheet1.getLastRowNum();
	for(int i=0; i<=lastRow; i++){
	Row row = sheet1.getRow(i);
	org.apache.poi.ss.usermodel.Cell cell = row.createCell(2);

	cell.setCellValue("WriteintoExcel");

	}

	FileOutputStream fos = new FileOutputStream(path);
	wb.write(fos);
	fos.close();
	}


}

