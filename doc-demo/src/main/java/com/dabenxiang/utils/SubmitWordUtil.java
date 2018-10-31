package com.dabenxiang.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Date:2018/10/13
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
public class SubmitWordUtil  {

    private String templatePath = "";
    private FileInputStream is = null;
    private XWPFDocument doc;

    private OutputStream os = null;


    public SubmitWordUtil(String templatePath) {
        this.templatePath = templatePath;

    }





    public void init() throws IOException {
        is = new FileInputStream(new File(this.templatePath));
        doc = new XWPFDocument(is);

    }



    public static void writeXWPFDocument(String[][] data){

        try {
            //创建一个word文档
            FileOutputStream outputStream  = new FileOutputStream("C:\\Users\\dabenxiang\\Desktop\\word1.docx");

            XWPFDocument xwpfDocument = new XWPFDocument();


            setParagraphTitle(xwpfDocument);


            /**
             * 创建一个table
             */
            //创建一个10行10列的表格
            XWPFTable table =xwpfDocument.createTable(data.length, data[0].length);
            setTableWidth(table,"8000");

            List<XWPFTableRow> rowList = table.getRows();


            for (int i = 0; i < rowList.size(); i++) {
                if (i == 0) {
                    setTitileRow(rowList.get(i));
                }else{
                    XWPFTableRow xwpfTableRow = rowList.get(i);
                    List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
                    for (int j = 0; j < tableCells.size(); j++) {
                        if(j==3){
                            setCellSizeAndGet(table,i,j,200,2000);
                        }else {
                            setCellSizeAndGet(table,i,j);

                        }
                    }

                }

            }


//            rowPr.

//            row.setHeight(2000);








            //为这一行增加一列
//            row.addNewTableCell();
//            //获取行属性
//            CTTrPr rowPr = row.getCtRow().addNewTrPr();
//            row.getCtRow();
//
//            /**
//             * 获取表格中的列  以及设计列样式
//             */
//            //获取某个单元格
//            XWPFTableCell cell ;
//            cell = row.getCell(0);
//            cell.setText("第一行\r\n第一列");
//            //单元格背景颜色
//            cell.setColor("676767");
//            //获取单元格样式
//            CTTcPr cellPr = cell.getCTTc().addNewTcPr();
//            //表格内容垂直居中
//            cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//            //设置单元格的宽度
//            cellPr.addNewTcW().setW(BigInteger.valueOf(5000));


            xwpfDocument.write(outputStream);
            outputStream.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }








    public static void main(String[] args) {
        String[][] table = new String[4][5];


        table[0][0] = "标题";
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
        table[2][2] = null;
        table[2][3] = null;
        table[2][4] = null;


        table[3][0] = "abc";
        table[3][1] = "abc";
        table[3][2] = null;
        table[3][3] = null;
        table[3][4] = null;


        writeXWPFDocument(table);
    }


    /**
     * 设置整个表格居中
     * @param table
     * @param width
     */
    private static  void setTableWidth(XWPFTable table,String width){
        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        CTJc cTJc=tblPr.addNewJc();
        cTJc.setVal(STJc.Enum.forString("center"));
//        tblWidth.setW(new BigInteger(width));
        tblWidth.setType(STTblWidth.DXA);
    }


    /**
     * 标题的段落样式和值
     *
     * @return
     */
    public static  void setParagraphTitle( XWPFDocument xwpfDocument ){
        //标题
        XWPFParagraph titleMes = xwpfDocument.createParagraph();
        titleMes.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun r1 = titleMes.createRun();
        r1.setBold(true);
        r1.setFontFamily("微软雅黑");
        r1.setText("上级采用情况");//活动名称
        r1.setFontSize(12);

        XWPFParagraph second = xwpfDocument.createParagraph();
        second.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun r2 = second.createRun();
        r2.setBold(true);
        r2.setFontFamily("微软雅黑");
        r2.setText("                        时间范围  ");//活动名称
        r2.setFontSize(8);
        r2.setColor("777777");
//        r1.setBold(true);
    }


    /**
     * 设置通用的样式
     * @param row
     */

    public static void setCommonRow(XWPFTableRow row){
        List<XWPFTableCell> tableCells = row.getTableCells();
        for (XWPFTableCell tableCell : tableCells) {
            tableCell.setText("tttt");

        }
    }


    /**
     * 设置标题这一行的样式
     * @param row
     */
    public static void setTitileRow(XWPFTableRow row){
        List<XWPFTableCell> tableCells = row.getTableCells();
        for (XWPFTableCell tableCell : tableCells) {
            setTitleCell(tableCell);
        }
        
    }


    /**
     * 设置标题每个单元个的样式
     * @param cell
     */
    public static void setTitleCell(XWPFTableCell cell ){

        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph p = new XWPFParagraph(ctp, cell);
        p.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = p.createRun();
        run.setColor("000000");
        run.setFontSize(10);
        run.setText("abc");
        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii("微软雅黑");
        fonts.setEastAsia("微软雅黑");
        fonts.setHAnsi("微软雅黑");
        cell.setParagraph(p);
        cell.setColor("DDDDDD");

    }


    /**
     * 使用默认的配置来设置
     * @param xTable
     * @param rowNomber
     * @param cellNumber
     * @return
     */
    private static XWPFTableCell setCellSizeAndGet(XWPFTable xTable,int rowNomber,int cellNumber) {
        return setCellSizeAndGet(xTable,rowNomber,cellNumber,200,1000);
    }

    //设置表格高度
    private static XWPFTableCell setCellSizeAndGet(XWPFTable xTable,int rowNomber,int cellNumber
        ,int height,int width){
        BigInteger bigWidth = new BigInteger(width+"");

        XWPFTableRow row = null;
        row =  xTable.getRow(rowNomber);
        row.setHeight(height);
        XWPFTableCell cell = null;
        cell = row.getCell(cellNumber);
         cell.getCTTc().addNewTcPr().addNewTcW().setW(bigWidth);
        return cell;
    }




}
