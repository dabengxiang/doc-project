package com.dabenxiang.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
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





}
