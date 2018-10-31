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
 * project name : doc-project
 * Date:2018/8/23
 * Author: yc.guo
 * DESC:
 */
public class WordDemo {

    private String templatePath = "";
    private FileInputStream is = null;
    private XWPFDocument doc;

    private OutputStream os = null;


    public WordDemo(String templatePath) {
        this.templatePath = templatePath;

    }


    public XWPFDocument getDoc() {
        return doc;
    }

    public void init() throws IOException {
        is = new FileInputStream(new File(this.templatePath));
        doc = new XWPFDocument(is);

    }


    /**
     * 替换掉占位符
     * @param params
     * @return
     * @throws Exception
     */
    public boolean export(Map<String,Object> params) throws Exception{
        this.replaceInPara(doc, params);
        return true;
    }


    /**
     * 替换文档里面的变量
     *
     * @param doc
     *            要替换的文档
     * @param params
     *            参数
     * @throws Exception
     */
    private void replaceInPara(XWPFDocument doc, Map<String, Object> params) throws Exception {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para = null;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }


    /**
     * 替换段落里面的变量
     *
     * @param para
     *            要替换的段落
     * @param params
     *            参数
     * @throws Exception
     * @throws IOException
     * @throws InvalidFormatException
     */
    private boolean replaceInPara(XWPFParagraph para, Map<String, Object> params) throws Exception {
        boolean data = false;
        List<XWPFRun> runs;
        //有符合条件的占位符
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            data = true;
            Map<Integer,String> tempMap = new HashMap<Integer,String>();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                //以"$"开头
                boolean begin = runText.indexOf("$")>-1;
                boolean end = runText.indexOf("}")>-1;
                if(begin && end){
                    tempMap.put(i, runText);
                    fillBlock(para, params, tempMap, i);
                    continue;
                }else if(begin && !end){
                    tempMap.put(i, runText);
                    continue;
                }else if(!begin && end){
                    tempMap.put(i, runText);
                    fillBlock(para, params, tempMap, i);
                    continue;
                }else{
                    if(tempMap.size()>0){
                        tempMap.put(i, runText);
                        continue;
                    }
                    continue;
                }
            }
        } else if (this.matcherRow(para.getParagraphText())) {
            runs = para.getRuns();
            data = true;
        }
        return data;
    }


    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
                Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }


    /**
     * 填充run内容
     * @param para
     * @param params
     * @param tempMap
     * @throws InvalidFormatException
     * @throws IOException
     * @throws Exception
     */
    private void fillBlock(XWPFParagraph para, Map<String, Object> params,
                           Map<Integer, String> tempMap, int index)
            throws InvalidFormatException, IOException, Exception {
        Matcher matcher;
        if(tempMap!=null&&tempMap.size()>0){
            String wholeText = "";
            List<Integer> tempIndexList = new ArrayList<Integer>();
            for(Map.Entry<Integer, String> entry :tempMap.entrySet()){
                tempIndexList.add(entry.getKey());
                wholeText+=entry.getValue();
            }
            if(wholeText.equals("")){
                return;
            }
            matcher = this.matcher(wholeText);
            if (matcher.find()) {
                boolean isPic = false;
                int width = 0;
                int height = 0;
                int picType = 0;
                String path = null;
                String keyText = matcher.group().substring(2,matcher.group().length()-1);
                Object value = params.get(keyText);
                String newRunText = "";
                if(value instanceof String){
                    newRunText = matcher.replaceFirst(String.valueOf(value));
                }else if(value instanceof Map){//插入图片
                    isPic = true;
                    Map pic = (Map)value;
                    width = Integer.parseInt(pic.get("width").toString());
                    height = Integer.parseInt(pic.get("height").toString());
                    picType = getPictureType(pic.get("type").toString());
                    path = pic.get("path").toString();
                }

                //模板样式				
                XWPFRun tempRun = null;
                // 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                // 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                for(Integer pos : tempIndexList){
                    tempRun = para.getRuns().get(pos);
                    tempRun.setText("", 0);
                }
                if(isPic){
                    //addPicture方法的最后两个参数必须用Units.toEMU转化一下
                    //para.insertNewRun(index).addPicture(getPicStream(path), picType, "测试",Units.toEMU(width), Units.toEMU(height));
                    tempRun.addPicture(getPicStream(path), picType, "测试",Units.toEMU(width), Units.toEMU(height));
                }else{
                    //样式继承
                    if(newRunText.indexOf("\n")>-1){
                        String[] textArr = newRunText.split("\n");
                        if(textArr.length>0){
                            //设置字体信息
                            String fontFamily = tempRun.getFontFamily();
                            int fontSize = tempRun.getFontSize();
                            //logger.info("------------------"+fontSize);
                            for(int i=0;i<textArr.length;i++){
                                if(i==0){
                                    tempRun.setText(textArr[0],0);
                                }else{
                                    if(StringUtils.isNotEmpty(textArr[i])){
                                        XWPFRun newRun=para.createRun();
                                        //设置新的run的字体信息
                                        newRun.setFontFamily(fontFamily);
                                        if(fontSize==-1){
                                            newRun.setFontSize(10);
                                        }else{
                                            newRun.setFontSize(fontSize);
                                        }
                                        newRun.addBreak();
                                        newRun.setText(textArr[i], 0);
                                    }
                                }
                            }
                        }
                    }else{
                        tempRun.setText(newRunText,0);
                    }
                }
            }
            tempMap.clear();
        }
    }


    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private boolean matcherRow(String str) {
        Pattern pattern = Pattern.compile("\\$\\[(.+?)\\]",
                Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher.find();
    }


    /**
     * 根据图片类型，取得对应的图片类型代码 
     * @param picType
     * @return int
     */
    private int getPictureType(String picType){
        int res = XWPFDocument.PICTURE_TYPE_PICT;
        if(picType != null){
            if(picType.equalsIgnoreCase("png")){
                res = XWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = XWPFDocument.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = XWPFDocument.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = XWPFDocument.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = XWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }



    private InputStream getPicStream(String picPath) throws Exception{
        URL url = new URL(picPath);
        //打开链接  
        HttpURLConnection conn = (HttpURLConnection)url.openConnection();
        //设置请求方式为"GET"  
        conn.setRequestMethod("GET");
        //超时响应时间为5秒  
        conn.setConnectTimeout(5 * 1000);
        //通过输入流获取图片数据  
        InputStream is = conn.getInputStream();
        return is;
    }

    public boolean generate(String outDocPath) throws IOException{
        os = new FileOutputStream(outDocPath);
        doc.write(os);
        this.close(os);
        this.close(is);
        return true;
    }

    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }



    /**
     * 关闭输出流
     *
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }



    /**
     * 替换掉表格中的占位符
     * @param params
     * @param tableIndex
     * @return
     * @throws Exception
     */
    public boolean export(Map<String,Object> params,int tableIndex) throws Exception{
        this.replaceInTable(doc, params,tableIndex);
        return true;
    }


    /**
     * 替换表格里面的变量
     *
     * @param doc
     *            要替换的文档
     * @param params
     *            参数
     * @throws Exception
     */
    private void replaceInTable(XWPFDocument doc, Map<String, Object> params,int tableIndex) throws Exception {
        List<XWPFTable> tableList = doc.getTables();
        if(tableList.size()<=tableIndex){
            throw new Exception("tableIndex对应的表格不存在");
        }
        XWPFTable table = tableList.get(tableIndex);
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        rows = table.getRows();
        for (XWPFTableRow row : rows) {
            cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                paras = cell.getParagraphs();
                for (XWPFParagraph para : paras) {
                    this.replaceInPara(para, params);
                }
            }
        }
    }


    /**
     * 循环生成表格
     * @param params
     * @param tableIndex
     * @return
     * @throws Exception
     */
    public boolean export(List<Map<String,String>> params,int tableIndex) throws Exception{

        return export(params,tableIndex,false);
    }

    public boolean export(List<Map<String,String>> params,int tableIndex,Boolean hasTotalRow) throws Exception{
        this.insertValueToTable(doc, params, tableIndex,hasTotalRow);
        return true;
    }



    private void insertValueToTable(XWPFDocument doc,List<Map<String, String>> params,int tableIndex) throws Exception {
        insertValueToTable(doc,params,tableIndex,false);
    }

    private void insertValueToTable(XWPFDocument doc,List<Map<String, String>> params,int tableIndex,Boolean hasTotalRow) throws Exception {
        List<XWPFTable> tableList = doc.getTables();
        if(tableList.size()<=tableIndex){
            throw new Exception("tableIndex对应的表格不存在");
        }
        XWPFTable table = tableList.get(tableIndex);
        List<XWPFTableRow> rows = table.getRows();
        if(rows.size()<2){
            throw new Exception("tableIndex对应表格应该为2行");
        }
        //模板行
        XWPFTableRow tmpRow = rows.get(1);
        List<XWPFTableCell> tmpCells = null;
        List<XWPFTableCell> cells = null;
        XWPFTableCell tmpCell = null;

        tmpCells = tmpRow.getTableCells();
        String cellText = null;
        String cellTextKey = null;

        Map<String,Object> totalMap = null;

        for (int i = 0, len = params.size(); i < len; i++) {
            Map<String,String> map = params.get(i);
            if(map.containsKey("total")){
                totalMap = new HashMap<String,Object>();
                totalMap.put("total", map.get("total"));
                continue;
            }
            XWPFTableRow row = table.createRow();
            row.setHeight(tmpRow.getHeight());
            cells = row.getTableCells();
            // 插入的行会填充与表格第一行相同的列数
            for (int k = 0, klen = cells.size(); k < klen; k++) {
                tmpCell = tmpCells.get(k);
                XWPFTableCell cell = cells.get(k);
                cellText = tmpCell.getText();
                if(StringUtils.isNotBlank(cellText)){
                    //转换为mapkey对应的字段
                    cellTextKey = cellText.replace("$", "").replace("{", "").replace("}","");
                    if(map.containsKey(cellTextKey)){
                        setCellText(tmpCell, cell, map.get(cellTextKey));
                    }
                }
            }
        }
        // 删除模版行
        table.removeRow(1);
        if(hasTotalRow && totalMap!=null){
            XWPFTableRow row = table.getRow(1);
            //cell.setText("11111");
            XWPFTableCell cell = row.getCell(0);
            replaceInPara(cell.getParagraphs().get(0),totalMap);

			/*String wholeText = cell.getText();
			Matcher matcher = this.matcher(wholeText);
			if(matcher.find()){*/
			/*List<XWPFParagraph> paras = cell.getParagraphs();

			for (XWPFParagraph para : paras) {
				this.replaceInPara(para, totalMap);
			}*/
            //}
            table.addRow(row);
            table.removeRow(1);
        }
    }


    private void setCellText(XWPFTableCell tmpCell, XWPFTableCell cell,
                             String text) throws Exception {

        CTTc cttc2 = tmpCell.getCTTc();
        CTTcPr ctPr2 = cttc2.getTcPr();

        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        //cell.setColor(tmpCell.getColor());
        // cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
        if (ctPr2.getTcW() != null) {
            ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
        }
        if (ctPr2.getVAlign() != null) {
            ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
        }
        if (cttc2.getPList().size() > 0) {
            CTP ctp = cttc2.getPList().get(0);
            if (ctp.getPPr() != null) {
                if (ctp.getPPr().getJc() != null) {
                    cttc.getPList().get(0).addNewPPr().addNewJc()
                            .setVal(ctp.getPPr().getJc().getVal());
                }
            }
        }

        if (ctPr2.getTcBorders() != null) {
            ctPr.setTcBorders(ctPr2.getTcBorders());
        }

        XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
        XWPFParagraph cellP = cell.getParagraphs().get(0);
        XWPFRun tmpR = null;
        if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
            tmpR = tmpP.getRuns().get(0);
        }

        List<XWPFRun> runList = new ArrayList<XWPFRun>();
        if(text==null){
            XWPFRun cellR = cellP.createRun();
            runList.add(cellR);
            cellR.setText("");
        }else{
            //这里的处理思路是：$b认为是段落的分隔符，分隔后第一个段落认为是要加粗的
            if(text.indexOf("\b") > -1){//段落，加粗，主要用于产品行程
                String[] bArr = text.split("\b");
                for(int b=0;b<bArr.length;b++){
                    XWPFRun cellR = cellP.createRun();
                    runList.add(cellR);
                    if(b==0){//默认第一个段落加粗
                        cellR.setBold(true);
                    }
                    if(bArr[b].indexOf("\n") > -1){
                        String[] arr = bArr[b].split("\n");
                        for(int i = 0; i < arr.length; i++){
                            if(i > 0){
                                cellR.addBreak();
                            }
                            cellR.setText(arr[i]);
                        }
                    }else{
                        cellR.setText(bArr[b]);
                    }
                }
            }else{
                XWPFRun cellR = cellP.createRun();
                runList.add(cellR);
                if(text.indexOf("\n") > -1){
                    String[] arr = text.split("\n");
                    for(int i = 0; i < arr.length; i++){
                        if(i > 0){
                            cellR.addBreak();
                        }
                        cellR.setText(arr[i]);
                    }
                }else{
                    cellR.setText(text);
                }
            }

        }

        // 复制字体信息
        if (tmpR != null) {
            //cellR.setBold(tmpR.isBold());
            //cellR.setBold(true);
            for(XWPFRun cellR : runList){
                if(!cellR.isBold()){
                    cellR.setBold(tmpR.isBold());
                }
                cellR.setItalic(tmpR.isItalic());
                cellR.setStrike(tmpR.isStrike());
                cellR.setUnderline(tmpR.getUnderline());
                cellR.setColor(tmpR.getColor());
                cellR.setTextPosition(tmpR.getTextPosition());
                if (tmpR.getFontSize() != -1) {
                    cellR.setFontSize(tmpR.getFontSize());
                }
                if (tmpR.getFontFamily() != null) {
                    cellR.setFontFamily(tmpR.getFontFamily());
                }
                if (tmpR.getCTR() != null) {
                    if (tmpR.getCTR().isSetRPr()) {
                        CTRPr tmpRPr = tmpR.getCTR().getRPr();
                        if (tmpRPr.isSetRFonts()) {
                            CTFonts tmpFonts = tmpRPr.getRFonts();
                            CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR
                                    .getCTR().getRPr() : cellR.getCTR().addNewRPr();
                            CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr
                                    .getRFonts() : cellRPr.addNewRFonts();
                            cellFonts.setAscii(tmpFonts.getAscii());
                            cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
                            cellFonts.setCs(tmpFonts.getCs());
                            cellFonts.setCstheme(tmpFonts.getCstheme());
                            cellFonts.setEastAsia(tmpFonts.getEastAsia());
                            cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
                            cellFonts.setHAnsi(tmpFonts.getHAnsi());
                            cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
                        }
                    }
                }
            }
        }
        // 复制段落信息
        cellP.setAlignment(tmpP.getAlignment());
        cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
        cellP.setBorderBetween(tmpP.getBorderBetween());
        cellP.setBorderBottom(tmpP.getBorderBottom());
        cellP.setBorderLeft(tmpP.getBorderLeft());
        cellP.setBorderRight(tmpP.getBorderRight());
        cellP.setBorderTop(tmpP.getBorderTop());
        cellP.setPageBreak(tmpP.isPageBreak());
        if (tmpP.getCTP() != null) {
            if (tmpP.getCTP().getPPr() != null) {
                CTPPr tmpPPr = tmpP.getCTP().getPPr();
                CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP
                        .getCTP().getPPr() : cellP.getCTP().addNewPPr();
                // 复制段落间距信息
                CTSpacing tmpSpacing = tmpPPr.getSpacing();
                if (tmpSpacing != null) {
                    CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr
                            .getSpacing() : cellPPr.addNewSpacing();
                    if (tmpSpacing.getAfter() != null) {
                        cellSpacing.setAfter(tmpSpacing.getAfter());
                    }
                    if (tmpSpacing.getAfterAutospacing() != null) {
                        cellSpacing.setAfterAutospacing(tmpSpacing
                                .getAfterAutospacing());
                    }
                    if (tmpSpacing.getAfterLines() != null) {
                        cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
                    }
                    if (tmpSpacing.getBefore() != null) {
                        cellSpacing.setBefore(tmpSpacing.getBefore());
                    }
                    if (tmpSpacing.getBeforeAutospacing() != null) {
                        cellSpacing.setBeforeAutospacing(tmpSpacing
                                .getBeforeAutospacing());
                    }
                    if (tmpSpacing.getBeforeLines() != null) {
                        cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
                    }
                    if (tmpSpacing.getLine() != null) {
                        cellSpacing.setLine(tmpSpacing.getLine());
                    }
                    if (tmpSpacing.getLineRule() != null) {
                        cellSpacing.setLineRule(tmpSpacing.getLineRule());
                    }
                }
                // 复制段落缩进信息
                CTInd tmpInd = tmpPPr.getInd();
                if (tmpInd != null) {
                    CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd()
                            : cellPPr.addNewInd();
                    if (tmpInd.getFirstLine() != null) {
                        cellInd.setFirstLine(tmpInd.getFirstLine());
                    }
                    if (tmpInd.getFirstLineChars() != null) {
                        cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
                    }
                    if (tmpInd.getHanging() != null) {
                        cellInd.setHanging(tmpInd.getHanging());
                    }
                    if (tmpInd.getHangingChars() != null) {
                        cellInd.setHangingChars(tmpInd.getHangingChars());
                    }
                    if (tmpInd.getLeft() != null) {
                        cellInd.setLeft(tmpInd.getLeft());
                    }
                    if (tmpInd.getLeftChars() != null) {
                        cellInd.setLeftChars(tmpInd.getLeftChars());
                    }
                    if (tmpInd.getRight() != null) {
                        cellInd.setRight(tmpInd.getRight());
                    }
                    if (tmpInd.getRightChars() != null) {
                        cellInd.setRightChars(tmpInd.getRightChars());
                    }
                }
            }
        }
    }


    public static void writeXWPFDocument(){

        try {
            //创建一个word文档
            XWPFDocument xwpfDocument = new XWPFDocument();
            FileOutputStream outputStream  = new FileOutputStream("C:\\Users\\dabenxiang\\Desktop\\word1.docx");
            /**
             * 创建一个段落
             */
            XWPFParagraph paragraph = xwpfDocument.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("德玛西亚！！");
            //加粗
            run.setBold(true);

            run = paragraph.createRun();
            run.setText("艾欧尼亚");
            run.setColor("fff000");


            /**
             * 创建一个table
             */
            //创建一个10行10列的表格
            XWPFTable table =xwpfDocument.createTable(10, 10);
//            table.createRow();
//            table.createRow();
            //添加新的一列
            table.addNewCol();
            //添加新的一行
//            table.createRow();
            //获取表格属性
            CTTblPr tablePr = table.getCTTbl().addNewTblPr();
            //获取表格宽度
            CTTblWidth tableWidth = tablePr.addNewTblW();
            //设置表格的宽度大小
            tableWidth.setW(BigInteger.valueOf(8000));

            /**
             * 获取表格中的行  以及设计行样式
             */
            //获取表格中的所有行
            List<XWPFTableRow> rowList = table.getRows();
            XWPFTableRow row;
            row = rowList.get(0);
            row.setHeight(2000);
            //为这一行增加一列
            row.addNewTableCell();
            //获取行属性
            CTTrPr rowPr = row.getCtRow().addNewTrPr();
            row.getCtRow();

            /**
             * 获取表格中的列  以及设计列样式
             */
            //获取某个单元格
            XWPFTableCell cell ;
            cell = row.getCell(0);
            cell.setText("第一行\r\n第一列");
            //单元格背景颜色
            cell.setColor("676767");
            //获取单元格样式
            CTTcPr cellPr = cell.getCTTc().addNewTcPr();
            //表格内容垂直居中
            cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
            //设置单元格的宽度
            cellPr.addNewTcW().setW(BigInteger.valueOf(5000));


            xwpfDocument.write(outputStream);
            outputStream.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        writeXWPFDocument();
    }





    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table     需要插入数据的表格
     * @param tableData 插入数据集合
     */
    private static void insertTable(XWPFTable table, String[][] tableData) {
        //创建行,根据需要插入的数据添加新行，不处理表头
        //遍历表格插入数据
        int offset = 1, tblIdx = 0;
        for (int i = 0; i < tableData.length; i++) {
            tblIdx = i + offset;
            XWPFTableRow tblRow = null;
            if (tblIdx >= table.getRows().size()) {
                tblRow = table.createRow();
            } else {
                tblRow = table.getRow(tblIdx);
            }

            List<XWPFTableCell> cells = tblRow.getTableCells();
            for (int j = 0; j < cells.size() && j<tableData[i].length; j++) {
                XWPFTableCell cell = cells.get(j);
                XWPFParagraph para = cell.getParagraphs().get(0);
                XWPFRun run = para.createRun();
                String s = tableData[i][j];
                //获取单元格样式
                CTTcPr cellPr = cell.getCTTc().addNewTcPr();
                //表格内容垂直居中
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);

                if(s.indexOf("\n") > -1) {
                    String[] text = s.split("\n");
                    for(int f = 0; f <= text.length -1; f++ ){
                        run.setText(text[f]);
                        if(text.length > 1 && text.length-1 != f) {
                            run.addBreak();
                        }
                    }
                }else {
                    run.setText(s);
                }
            }
        }
    }

    /**
     * 替换表格里面的变量
     * @param params 参数
     * @param tableDatas
     */
    public void replaceInTable( Map<String, Object> params, String[][][] tableDatas) throws Exception {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        int idx = 0;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext() && idx < tableDatas.length) {
            table = iterator.next();
            if (table.getRows().size() >= 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if (this.matcher(table.getText()).find()) {
                    rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {
                            paras = cell.getParagraphs();
                            for (XWPFParagraph para : paras) {
                                this.replaceInPara(para, params);
                            }
                        }
                    }
                } else {
                    insertTable(table, tableDatas[idx]);  //插入数据
                }
            }
        }
    }




}
