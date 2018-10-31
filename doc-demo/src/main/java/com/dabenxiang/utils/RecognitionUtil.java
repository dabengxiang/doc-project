package com.dabenxiang.utils;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

/**
 * Date:2018/10/31
 * Author: yc.guo the one whom in nengxun
 * Desc:
 */
public class RecognitionUtil {


    public static void main(String[] args) throws IOException {
        String templatePath = "C:\\Users\\dabenxiang\\Desktop\\3.docx";
        FileInputStream is = new FileInputStream(new File(templatePath));
        XWPFDocument doc = new XWPFDocument(is);
        getWordText(doc);
    }


    public  static void getWordText(XWPFDocument doc){

        List<XWPFPictureData> allPictures = doc.getAllPictures();



        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para = null;

        while (iterator.hasNext()) {
            para = iterator.next();
            List<XWPFRun> runs;
            runs = para.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                System.out.println(runText);
            }
        }

    }

}
