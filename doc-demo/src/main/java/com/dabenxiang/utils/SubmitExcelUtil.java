package com.dabenxiang.utils;

import com.dabenxiang.properties.MergeProperties;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.*;





/**
 * Date:2018/10/13
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
public class SubmitExcelUtil {

    public static boolean createExcel(String targetFilePath, String[][] table, List<MergeProperties> mergeList) {
        Workbook wb = new XSSFWorkbook();
        Font titleFont = wb.createFont();
        CellStyle titleStyle = getDefaultCellStyle(wb);
//        CellStyle valueCellStyle = getDefaultCellStyle(wb);
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
        sheet.setColumnWidth(1,30 * 256);
        sheet.setColumnWidth(2,50 * 256);
        sheet.setColumnWidth(3,30 * 256);
        sheet.setColumnWidth(4,10 * 256);


        for (MergeProperties mergeProperties : mergeList) {
            CellRangeAddress cra1 =new CellRangeAddress(mergeProperties.getStartRow(),
                    mergeProperties.getEndRow(), mergeProperties.getStartCol(), mergeProperties.getEndCol());

            sheet.addMergedRegion(cra1);
        }

//
//        sheet.addMergedRegion(cra1);
//        sheet.addMergedRegion(cra2);
//        sheet.addMergedRegion(cra3);


        for (int i = 0; i < table.length; i++) {
            Row row = sheet.createRow(i);
            row.setHeightInPoints(40);

            boolean flag = false;

            for (int j = 0; j < 5; j++) {
                Cell cell = row.createCell(j);

                cell.setCellStyle(keyCellStyle);


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


    public  static CellStyle getDefaultCellStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;


    }


    // 测试
    public static void main(String[] args) {


        String[][] table = new String[5][5];


        table[0][0] = "上级采用情况";
        table[0][1] = null;
        table[0][2] = null;
        table[0][3] = null;
        table[0][4] = null;


        table[1][0] = "abc";
        table[1][1] = "abc";
        table[1][2] = "abc";
        table[1][3] = "abc";
        table[1][4] = "abc";



        table[2][0] = "abc";
        table[2][1] = "abc";
        table[2][2] = "abc";
        table[2][3] = "abc";
        table[2][4] = "abc";

        table[3][0] = "abc";
        table[3][1] = "abc";
        table[3][2] = "abc";
        table[3][3] = "abc";
        table[3][4] = "abc";


        table[4][0] = "abc";
        table[4][1] = "abc";
        table[4][2] = "abc";
        table[4][3] = "abc";
        table[4][4] = "abc";




//        List<String> list = new ArrayList<>();
//
//        list.add("标题");
//        list.add("456");
//        list.add("789");
//        list.add("111");
//        list.add("222");
//        list.add("3333");
//        list.add("4444");
//
//
//        Map item = new HashMap();
//        item.put("abc","优秀1213");
//        item.put("L-00002","L-00013");
//        item.put("L-00003","L-00014");
//
//        String path =  "C:\\Users\\dabenxiang\\Desktop\\导出单个活动的数据.xlsx";
        String path2 =  "C:\\Users\\dabenxiang\\Desktop\\1.xlsx";
//        replaceModel(item, path, path2);

        List<MergeProperties> mergeList = new ArrayList<>();

        MergeProperties mergeProperties1 = new MergeProperties(2, 3, 1, 1);

        MergeProperties mergeProperties2 = new MergeProperties(0, 0, 0, table[0].length-1);

        mergeList.add(mergeProperties1);
        mergeList.add(mergeProperties2);

        createExcel(path2,table,mergeList);

    }

}
