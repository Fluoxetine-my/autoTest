package com.test.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {
    public static Object[][] getData(String filePath, String fileName, String sheetName) throws IOException {
        File file = new File(filePath + "\\" + fileName);
        //创建FileInputStream对象用于读取Excel文件
        FileInputStream inputStream = new FileInputStream(file);
        //声明Workbook对象
        Workbook workbook = null;
        //获取文件名参数的扩展名，判断是.xlsx文件还是.xls文件
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
        //如果是.xlsx,则用XSSFWorkbook对象进行实例化，如果是.xls则使用HSSFWorkbook对象进行实例化
        if (fileExtensionName.equals(".xlsx")){
            workbook = new XSSFWorkbook(inputStream);
        }
        else if (fileExtensionName.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }
        //通过sheetName参数生成Sheet对象
        Sheet sheet = workbook.getSheet(sheetName);
        //获取Excel数据文件Sheet1中数据的行数，getLastRowNum方法获取数据的最后一行行号
        //getFirstRowNum方法获取数据的第一行行号，相减之后算出数据的行数
        //Excel行和列都是从0开始
        int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
        //创建名为records的list对象来存储从Excel数据文件读取的数据
        List<Object[]> records = new ArrayList<Object[]>();
        //使用两个for循环遍历Excel数据文件除去第一行外所有数据
        //所以i从1开始，而不是从0开始
        for (int i = 1; i<rowCount+1; i++){
            Row row = sheet.getRow(i);
            //声明一个数组，用来存储Excel数据文件每行中的数据，数组的大小用getLastCellNum办法来进行动态声明，实现测试数据个数和数组大小相一致
            String fields[] = new String[row.getLastCellNum()];
            for (int j = 0; j<row.getLastCellNum();j++){
                //调用getCell和getStringCellValue方法获取Excel文件中的单元格数据
                fields[j] = row.getCell(j).getStringCellValue();
            }
            //将fields的数据兑现存储到records的list中
            records.add(fields);
        }
        //定义函数返回值，即Object[][]
        //将存储测试数据的list转换为一个Object的二维数组
        Object[][] results = new Object[records.size()][];
        //设置二维数组每行的值，每行是一个Object对象
        for (int i = 0; i<records.size(); i++){
            results[i] = records.get(i);
        }
        return results;
    }
}
