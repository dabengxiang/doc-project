package com.dabenxiang.mian;

import com.dabenxiang.utils.WordDemo;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * project name : doc-project
 * Date:2018/8/23
 * Author: yc.guo
 * DESC:
 */
public class DocMain {
        public  static void main(String[] args) throws Exception {
            exprotTableData();
        }



        public void exprotPara() throws Exception {
            String url = "C:\\Users\\dabenxiang\\Desktop\\test1.docx";
            WordDemo wordDemo = new WordDemo(url);
            wordDemo.init();
            HashMap<String, Object> map = new HashMap<>();
            map.put("abc","i yuyu ye");
            wordDemo.export(map);
            wordDemo.generate(url);

        }

        public static void exprotTable() throws Exception {
            String url = "C:\\Users\\dabenxiang\\Desktop\\test1.docx";
            WordDemo wordDemo = new WordDemo(url);
            wordDemo.init();
            List<Object> list = new ArrayList<>();
            HashMap<String, Object> map = new HashMap<>();
            map.put("abc","办公厅");
            wordDemo.export(map);
            wordDemo.export(map,0);
            wordDemo.generate(url);
        }


    public static void exprotTableData() throws Exception {
        String url = "C:\\Users\\dabenxiang\\Desktop\\abc.docx";
        WordDemo wordDemo = new WordDemo(url);
        wordDemo.init();
        List<Object> list = new ArrayList<>();
        HashMap<String, Object> map = new HashMap<>();
        String[][][] tableDatas = new String[1][4][4];
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 4; j++) {
                tableDatas[0][i][j] = String.valueOf(i)+",\n"+String.valueOf(j);
            }

        }

        wordDemo.replaceInTable(null,tableDatas);
        
        wordDemo.generate(url);
    }


}
