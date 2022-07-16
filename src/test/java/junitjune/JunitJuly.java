package junitjune;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JunitJuly {
	public static void main(String[] args) throws Exception {
		File f=new File("C:\\Users\\Sarath Kumar\\Desktop\\Book1.xlsx");
		FileInputStream i=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(i);
		Sheet s = w.getSheet("DUMMY");
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		String sv = c.getStringCellValue();
		System.out.println(sv);
		File f1=new File("C:\\Users\\Sarath Kumar\\Desktop\\Book1.xlsx");
		Workbook w1=new XSSFWorkbook();
		Sheet s1 = w1.createSheet("DUMMY");
		Row r1 = s1.createRow(0);
		Cell c1 = r1.createCell(0);
		c1.setCellValue("abc");
		FileOutputStream o=new FileOutputStream(f1);
		w1.write(o);
		
	}

}
