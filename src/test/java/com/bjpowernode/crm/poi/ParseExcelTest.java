package com.bjpowernode.crm.poi;

import com.bjpowernode.crm.commons.utils.HSSFUtils;
import com.bjpowernode.crm.commons.utils.UUIDUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * 使用apache-poi解析excel文件
 */
public class ParseExcelTest {
    public static void main(String[] args) throws IOException {
        //根据指定的excel文件生成HSSFWorkbook对象，封装了excel文件的所有信息
        InputStream is = new FileInputStream("D:\\JAVAPRO\\crm-project\\temp\\studentList.xls");
        HSSFWorkbook wb = new HSSFWorkbook(is);
        //根据wb获取HSSFSheet对象，封装了一页的所有信息
        HSSFSheet sheet = wb.getSheetAt(0);
        //根据sheet获取HSSFRow对象，封装了一行的所有信息
        HSSFRow row = null;
        HSSFCell cell = null;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) { //sheet.getLastRowNum() 获取最后一行的索引
            row = sheet.getRow(i);

            //根据row对象获取HSSFCell对象，封装了一列的所有信息
            for (int j = 0; j < row.getLastCellNum(); j++) { //row.getLastCellNum() 获取总列数 = 最后一行下标 + 1
                cell = row.getCell(j);

                //获取列中数据
               /* if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
                    System.out.print(cell.getBooleanCellValue() + " ");
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
                    System.out.print(cell.getCellFormula() + " ");
                } else {
                    System.out.print("" + " ");
                }*/
                System.out.print(HSSFUtils.getCellValueForStr(cell) + " ");
            }

            //每一行中所有列都打完，打印换行
            System.out.println();

        }
    }

    /**
     * 从指定的HSSFCell对象中获取列的值
     * @return
     */
    public static String getCellValueForStr(HSSFCell cell) {
        String ret = "";
        if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            ret = cell.getStringCellValue();
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            ret = cell.getNumericCellValue() + "";
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
            ret = cell.getBooleanCellValue() + "";
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            ret = cell.getCellFormula() + "";
        } else {
            ret = "";
        }
        return ret;
    }
}
