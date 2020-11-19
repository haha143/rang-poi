package com.rang.poi;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileOutputStream;

public class ExcelWrite {
    String PATH="D:/IDEA/projects/rang-poi/";
    @Test
    public void  testExcel03() throws Exception{
        //创建一个工作簿
        Workbook workbook=new HSSFWorkbook();
        //创建一张工作表
        Sheet sheet=workbook.createSheet("我是一个新表格");
        //创建一行即(1,1)的单元格
        Row row1=sheet.createRow(0);
        Cell cell11=row1.createCell(0);
        //往该单元格中填充数据
        cell11.setCellValue("姓名");
        //创建(1,2)单元格
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("印某人");


        Row row2=sheet.createRow(1);
        Cell cell21=row2.createCell(0);
        cell21.setCellValue("注册日期");
        Cell cell22=row2.createCell(1);
        String time=new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"登记表03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("文件生成成功");
    }
    @Test
    public void testExcel07()throws Exception{
        Workbook workbook=new XSSFWorkbook();
        Sheet sheet=workbook.createSheet("我是一个新表格");
        Row row1=sheet.createRow(0);
        Cell cell11=row1.createCell(0);
        cell11.setCellValue("姓名");
        Cell cell12=row1.createCell(1);
        cell12.setCellValue("印某人");


        Row row2=sheet.createRow(1);
        Cell cell21=row2.createCell(0);
        cell21.setCellValue("注册日期");
        Cell cell22=row2.createCell(1);
        String time=new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell22.setCellValue(time);

        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"登记表07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("文件生成成功");
    }

    @Test
    public void test03BigData()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new HSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<65536;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test03BigData.xls");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");

    }

    @Test
    public void test07BigData()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new XSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<100000;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test07BigData.xlsx");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");
    }

    @Test
    public void test07BigDataS()throws Exception{
        Long begin=System.currentTimeMillis();
        Workbook workbook=new SXSSFWorkbook();
        Sheet sheet=workbook.createSheet();
        for(int rownum=0;rownum<100000;rownum++){
            Row row=sheet.createRow(rownum);
            for(int cellnum=0;cellnum<10;cellnum++){
                Cell cell=row.createCell(cellnum);
                cell.setCellValue(cellnum);
            }
        }
        FileOutputStream fileOutputStream=new FileOutputStream(PATH+"test07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        System.out.println("文件生成完毕");
        Long end=System.currentTimeMillis();
        System.out.println("共用时:"+(double)(end-begin)/1000+"秒");
    }
}
