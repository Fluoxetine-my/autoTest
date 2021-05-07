package com.test.utils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelReport {
    public static int rowNumber = 1;

    public static void main(String[] args) {
        writeExcel("yuanqing","Login","testLoginFail01","注释","pass","");
    }

    /**
     *
     * @param packageName : 包名
     * @param className ： 类名
     * @param methodName ：方法名
     * @param remark ： 注释
     * @param result ： 结果 pass fail
     * @param reason ：原因 pass 则为空 ，fail则有失败原因
     */
    public static void writeExcel(String packageName , String className, String methodName , String remark , String result , String reason){
        try{
            /**
             * 可以尝试，每次都生成不同的excel文档，往里面添加内容，但是需要office excel
             */
            //report文件的路径
            String path = "src/test/java/com/test/datas/TestResult.xlsx" ;
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(path));
            XSSFSheet sheet=wb.getSheetAt(0);
            //获得EXCEL行数
            int rowNums=sheet.getLastRowNum();
            // System.out.println("多少行:" +rowNums);
            //往sheet中追加一行数据
            int rowCurrentNumber = rowNums+1;
            sheet.createRow(rowCurrentNumber);
            XSSFRow row = sheet.getRow(rowCurrentNumber);
            //格式
            CellStyle cellStyle2=wb.createCellStyle();
            cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex()); // 前景色
            cellStyle2.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cellStyle2.setBorderBottom(CellStyle.BORDER_THIN); // 底部边框
            if(row != null){
                //System.out.println("行不为空！" );
                Date now = new Date();
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");//可以任意地修改日期格式
                String currentTime = dateFormat.format( now );
                //创建单元格并赋值

                row.createCell(0).setCellValue(currentTime);
                row.createCell(1).setCellValue(packageName);
                row.createCell(2).setCellValue(className);
                row.createCell(3).setCellValue(methodName);
                row.createCell(4).setCellValue(remark);
                row.createCell(5).setCellValue(result);
                if(result.equals("fail")){
                    row.getCell(5).setCellStyle(cellStyle2);
                }
                row.createCell(6).setCellValue(reason);
            }else{
                //System.out.println("行为空！" );
            }
            FileOutputStream os = new FileOutputStream(path);
            wb.write(os);//一定要写这句代码，否则无法将数据写入excel文档中
            os.close();
        }catch(Exception e){
            e.printStackTrace();
        }

    }
}
