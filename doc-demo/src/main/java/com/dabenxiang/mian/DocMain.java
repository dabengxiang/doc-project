package com.dabenxiang.mian;

import com.dabenxiang.utils.WordDemo;

import java.io.IOException;
import java.util.HashMap;

/**
 * project name : doc-project
 * Date:2018/8/23
 * Author: yc.guo
 * DESC:
 */
public class DocMain {
        public  static void main(String[] args) throws Exception {
            String url = "C:\\Users\\83673\\Desktop\\test.docx";
            System.out.println(url);
            WordDemo wordDemo = new WordDemo(url);
            wordDemo.init();
            HashMap<String, Object> map = new HashMap<>();
            map.put("abc","i yuyu ye");
            wordDemo.export(map);
            wordDemo.generate(url);
            
        }
}
