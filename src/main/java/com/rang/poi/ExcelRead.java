package com.rang.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.FileInputStream;

public class ExcelRead {
    String PATH="D:/IDEA/projects/rang-poi/";

    @Test
    public void test03Read()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test03BigData.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1,Row.RETURN_NULL_AND_BLANK);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }

    @Test
    public void test07Read()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test07BigData.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1,Row.RETURN_NULL_AND_BLANK);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }

    @Test
    public void testMultipleTypeRead()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1,Row.RETURN_NULL_AND_BLANK);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }


}
