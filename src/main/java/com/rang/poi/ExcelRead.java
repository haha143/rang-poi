package com.rang.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import java.io.FileInputStream;

public class ExcelRead {
    String PATH="D:/software/IDEA/projects/rang-poi/";

    @Test
    public void test03Read()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test03BigData.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }

    @Test
    public void test07ReadByPolicy()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test07BigData.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
//        Cell cell=row.getCell(-1);
        Cell cell=row.getCell(-1,Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        System.out.println(cell);
    }


    @Test
    public void test07Read()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test07BigData.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(0);
        Cell cell[]=new Cell[4];
        cell[0]=row.getCell(0);
        cell[1]=row.getCell(1);
        cell[2]=row.getCell(2);
        cell[3]=row.getCell(3);
        for(int i=0;i<4;i++){
            System.out.println(cell[i]);
        }
    }

    @Test
    public void testMultipleTypeRead()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);
        Row rowTitle=sheet.getRow(0);
        int cellNum=rowTitle.getLastCellNum();
        if(rowTitle!=null){
            for(int i=0;i<cellNum;i++){
                Cell cell=rowTitle.getCell(i);
                int cellType=cell.getCellType();
                if(cell!=null){
                    System.out.print(cell+"-"+cellType+" | ");
                }
            }
        }
        System.out.println();
        int RowNum=sheet.getLastRowNum();
        for(int i=1;i<=RowNum;i++){
            Row rowData=sheet.getRow(i);
            if(rowData!=null){
                int cellnum=rowData.getLastCellNum();
                for(int j=0;j<cellnum;j++){
                    Cell cell=rowData.getCell(j);
                    int cellType=cell.getCellType();
                    if(cell!=null){
                        switch (cellType){
                            case HSSFCell.CELL_TYPE_NUMERIC:
                                System.out.print(cell.getNumericCellValue()+"-"+cellType+" | ");
                                continue;
                            case HSSFCell.CELL_TYPE_STRING:
                                System.out.print(cell.getStringCellValue()+"-"+cellType+" | ");
                                continue;
                            case HSSFCell.CELL_TYPE_FORMULA:
                                System.out.print("null"+"-"+cellType+" | ");
                                continue;
                            case HSSFCell.CELL_TYPE_BLANK:
                                System.out.print(cell.getStringCellValue()+"-"+cellType+" | ");
                                continue;
                            case HSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.print(cell.getBooleanCellValue()+"-"+cellType+" | ");
                                continue;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print(cell.getErrorCellValue()+"-"+cellType+" | ");
                                continue;
                        }
                    }
                }
                System.out.println();
            }
        }
    }


    @Test
    public void testFORMULA()throws Exception{
        FileInputStream fileInputStream=new FileInputStream(PATH+"test.xls");
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //获取到包含公式的单元格
        Sheet sheet=workbook.getSheetAt(0);
        Row row=sheet.getRow(3);
        Cell cell=row.getCell(7);
        //读取计算的公式
        FormulaEvaluator formulaEvaluator=new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
        int cellType=cell.getCellType();
        switch (cellType){
            //单元格的类型是公式类型
            case HSSFCell.CELL_TYPE_FORMULA:
                //公式内容
                String formula=cell.getCellFormula();
                System.out.println(formula);
                //执行公式之后,单元格内的值
                CellValue evaluate=formulaEvaluator.evaluate(cell);
                String cellValue=evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }

    }


}
