package com.poi;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;


public class ExcelWrite {

    //这道程序，需要先在本地新建文件夹才可以使用
    //两者之间除了输出本地文件的格式03是xls,07是xlsx，JAVA里面的对象不一样以外
    private String PATH="D:\\AHope\\";

    @Test
    public void write03() throws IOException {
        // 1、创建工作薄
        Workbook workbook = new HSSFWorkbook();
        //2、创建工作表
        Sheet sheet = workbook.createSheet("test1");

        //3、创建行(0:第一行)
        Row row1 = sheet.createRow(0);
        //4、创建单元格(0:第一行的第一个格子)
        Cell cell00 = row1.createCell(0);
        cell00.setCellValue("测试数据");
        //第一行的第二个格子
        Cell cell01 = row1.createCell(1);
        cell01.setCellValue("testData");

        //第二行
        Row row2 = sheet.createRow(1);
        //第二行的第一个格子
        Cell cell10 = row2.createCell(0);
        cell10.setCellValue("时间");
        //第二行的第二个格子
        Cell cell11 = row2.createCell(1);
        cell11.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生成Excel（IO流）
        //03版本的Excel是使用 .xls 结尾！！！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "POI03测试.xls");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("03Excel输出完毕");
    }

    @Test
    public void write07() throws IOException {
        // 1、创建工作薄
        Workbook workbook = new XSSFWorkbook();
        //2、创建工作表
        Sheet sheet = workbook.createSheet("test1");

        //3、创建行(0:第一行)
        Row row1 = sheet.createRow(0);
        //4、创建单元格(0:第一行的第一个格子)
        Cell cell00 = row1.createCell(0);
        cell00.setCellValue("测试数据");
        //第一行的第二个格子
        Cell cell01 = row1.createCell(1);
        cell01.setCellValue("testData");

        //第二行
        Row row2 = sheet.createRow(1);
        //第二行的第一个格子
        Cell cell10 = row2.createCell(0);
        cell10.setCellValue("时间");
        //第二行的第二个格子
        Cell cell11 = row2.createCell(1);
        cell11.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生成Excel（IO流）
        //03版本的Excel是使用 .xlsx 结尾！！！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "POI07测试.xlsx");

        workbook.write(fileOutputStream);

        //关闭流
        fileOutputStream.close();

        System.out.println("07Excel输出完毕");
    }

    @Test
    public void testwrite03BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();
        //创建一个薄
        Workbook workbook = new HSSFWorkbook();
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream fos = new FileOutputStream(PATH + "03版本Excel大量数据测试.xls");
        workbook.write(fos);
        fos.close();
        System.out.println("over");
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    @Test
    public void testwrite07_S_BigData() throws IOException {
        //时间
        long begin = System.currentTimeMillis();
        //创建一个薄
        Workbook workbook = new SXSSFWorkbook(100);
        //创建表
        Sheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        FileOutputStream fos = new FileOutputStream(PATH + "07_S_版本Excel大量数据测试.xlsx");
        workbook.write(fos);
        fos.close();
        //清除临时缓存
        ((SXSSFWorkbook)workbook).dispose();
        System.out.println("over");
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    @Test
    public void testRead03() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "POI03测试.xls");

        //1、获取工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        //2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3、得到行
        Row row = sheet.getRow(0);
        //4、得到列
        Cell cell = row.getCell(0);

        //读取值时一定要注意类型
        System.out.println(cell.getStringCellValue());

        fileInputStream.close();
    }

    @Test
    public void testRead07() throws Exception {
        //获取文件流
        FileInputStream fis = new FileInputStream(PATH + "POI07测试.xlsx");
        //1、创建一个工作簿。使用 exceL能操作的这边他都可以操作！
        Workbook workbook = new XSSFWorkbook(fis);
        //2、得到表
        Sheet sheet = workbook.getSheetAt(0);
        //3、得到行
        Row row = sheet.getRow(0);
        //4、得到列
        Cell cell = row.getCell(0);

        //读取值的时候，一定要注意类型！
        //getStringCellValue 字符串类型
        System.out.println(cell.getStringCellValue());
        fis.close();
    }

    @Test
    public void testCellType() throws Exception {
        //获取文件
        FileInputStream fileInputStream = new FileInputStream(PATH + "明细表.xls");

        //获取工作薄
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        //得到表
        Sheet sheet = workbook.getSheetAt(0);

        //获取标题内容
        Row rowTitle = sheet.getRow(0);
        if (rowTitle != null){
            //获取一行中有多少个单元格
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                //获取单元
                Cell cell = rowTitle.getCell(cellNum);
                if (cell != null){
                    //获取类型
                    int cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        //获取表中的内容
        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowNum = 1; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);
            if (rowData != null){
                //读取列
                int cellCout = rowTitle.getPhysicalNumberOfCells();
                for (int cellNum = 0; cellNum < cellCout; cellNum++) {
                    System.out.print("【" + (rowNum+1) + "-" + (cellNum+1) + "】");

                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if (cell != null){
                        int cellType = cell.getCellType();
                        String cellValue = "";

                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING://字符串
                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN://布尔值
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC://数字类型
                                System.out.print("【NUMERIC】");

                                if (HSSFDateUtil.isCellDateFormatted(cell)){//日期
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime().toString("yyyy-MM-dd");
                                }else{
                                    // 不是日期格式，则防止当数字过长时以科学计数法显示
                                    System.out.print("【转换成字符串】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_BLANK://空
                                System.out.print("【BLANK】");
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fileInputStream.close();
    }

    @Test
    public void testFormula() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "计算公式.xls");
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheetAt = workbook.getSheetAt(0);

        //获取计算公式所在的行
        Row row = sheetAt.getRow(4);
        //计算公式的第几个单元格
        Cell cell = row.getCell(0);

        //拿到计算公式
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);


        //输出单元格内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA://公式
                //获取单元格的计算公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }








}
