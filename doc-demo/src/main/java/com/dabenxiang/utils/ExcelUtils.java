package com.dabenxiang.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.lang.model.util.ElementScanner6;
import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Date:2018/9/17
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
public class ExcelUtils {

    /**
     * 替换Excel模板文件内容
     * @param item 文档数据
     * @param sourceFilePath Excel模板文件路径
     * @param targetFilePath Excel生成文件路径
     */
    public static boolean replaceModel(Map item, String sourceFilePath, String targetFilePath) {
        boolean bool = true;
        try {

            Workbook wb = null;

            if(StringUtils.endsWithIgnoreCase(sourceFilePath,".xls")){
                POIFSFileSystem fs  =new POIFSFileSystem(new FileInputStream(sourceFilePath));
                wb = new HSSFWorkbook(fs);
            }


            else if(StringUtils.endsWithIgnoreCase(sourceFilePath,".xlsx")){
                FileInputStream fs = new FileInputStream(sourceFilePath);
                wb = new XSSFWorkbook(fs);
            }

            else{
                throw  new RuntimeException("不支持此格式！");
            }


            Sheet sheet = wb.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while(rows.hasNext()){
                Row row = (Row) rows.next();
                if(row!=null) {
                    int num = row.getLastCellNum();
                    System.out.println(num);
                    for(int i=0;i<num;i++) {
                        Cell cell=  row.getCell(i);
                        if(cell!=null) {
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        }
                        if(cell==null || cell.getStringCellValue()==null) {
                            continue;
                        }
                        String value= cell.getStringCellValue().trim();
                        if(StringUtils.isNotBlank(value)) {
                            Set<String> keySet = item.keySet();
                            Iterator<String> it = keySet.iterator();
                            while (it.hasNext()) {
                                String text = it.next();
                                if(value.equalsIgnoreCase(text)) {
                                    cell.setCellValue((String)item.get(text));
                                    break;
                                }
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }

            // 输出文件
            FileOutputStream fileOut = new FileOutputStream(targetFilePath);
            wb.write(fileOut);
            fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;
    }

    public  static CellStyle getDefaultCellStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;


    }

    public static boolean createExcel(String targetFilePath,String[][] table) {
        Workbook wb = new XSSFWorkbook();
        Font titleFont = wb.createFont();
        CellStyle titleStyle = getDefaultCellStyle(wb);
        CellStyle valueCellStyle = getDefaultCellStyle(wb);
        CellStyle keyCellStyle = getDefaultCellStyle(wb);

        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        keyCellStyle.setAlignment(HorizontalAlignment.CENTER);
        keyCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);
        titleStyle.setFont(titleFont);

        Sheet sheet =  wb.createSheet();


        sheet.setDisplayGridlines(false);

        sheet.setColumnWidth(0,10 * 256);
        sheet.setColumnWidth(1,50 * 256);
        sheet.setColumnWidth(2,16 * 256);
        sheet.setColumnWidth(3,50 * 256);


//        CellRangeAddress cra1 =new CellRangeAddress(0, 0, 0, 3); // 起始行, 终止行, 起始列, 终止列
//        CellRangeAddress cra2 =new CellRangeAddress(3, 3, 1, 3); // 起始行, 终止行, 起始列, 终止列
//        CellRangeAddress cra3 =new CellRangeAddress(8, 8, 1, 3); // 起始行, 终止行, 起始列, 终止列
//
//
//        sheet.addMergedRegion(cra1);
//        sheet.addMergedRegion(cra2);
//        sheet.addMergedRegion(cra3);


        for (int i = 0; i < table.length; i++) {
            Row row = sheet.createRow(i);
            if(i==0){
                row.setHeightInPoints(50);
            }else {
                row.setHeightInPoints(40);

            }

            boolean flag = false;

            for (int j = 0; j < 4; j++) {
                Cell cell = row.createCell(j);
                if(j%2 == 0)
                    cell.setCellStyle(keyCellStyle);
                else
                    cell.setCellStyle(valueCellStyle);
                if(table[i][j] == null){
                    if(flag)
                        continue;
                    else{
                        CellRangeAddress cra =new CellRangeAddress(i, i, j-1, table[i].length-1); // 起始行, 终止行, 起始列, 终止列
                        sheet.addMergedRegion(cra);
                        flag = true;
                    }
                    continue;
                }
                if(i==0){
                    cell.setCellStyle(titleStyle);
                    cell.setCellValue(table[0][0]);
                    continue;
                }
                cell.setCellValue(table[i][j]);

            }
        }
        Iterator<Row> rows = sheet.rowIterator();

        while(rows.hasNext()){
            Row row = (Row) rows.next();
            if(row!=null) {
                int num = row.getLastCellNum();
                System.out.println(num);
            }
        }


        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(targetFilePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }



        return true;

    }

    // 测试
    public static void main(String[] args) {


        String[][] table = new String[3][4];


        table[0][0] = "标题";
        table[0][1] = null;
        table[0][2] = null;
        table[0][3] = null;

        table[1][0] = "abc";
        table[1][1] = "abc";
        table[1][2] = "abc";
        table[1][3] = "abc";


        table[2][0] = "abc";
        table[2][1] = "abc";
        table[2][2] = null;
        table[2][3] = null;


        List<String> list = new ArrayList<>();

        list.add("标题");
        list.add("456");
        list.add("789");
        list.add("111");
        list.add("222");
        list.add("3333");
        list.add("4444");


        Map item = new HashMap();
        item.put("abc","优秀1213");
        item.put("L-00002","L-00013");
        item.put("L-00003","L-00014");

        String path =  "C:\\Users\\dabenxiang\\Desktop\\导出单个活动的数据.xlsx";
        String path2 =  "C:\\Users\\dabenxiang\\Desktop\\1.xlsx";
        replaceModel(item, path, path2);
//        createExcel(path2,table);


    }

}
