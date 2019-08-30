package com.elens.convert.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.util.UUID;

public class JacobUtil {
    // 12、16 代表html保存成word(docx)
    public static final int HTML_WORD = 12;

    // 0 代表html保存成word(doc)
    //public static final int HTML_WORD = 0;

    public static void main(String[] args) {
        String wordFile = "C:\\Users\\yilanqunzhi\\Desktop\\qwe.doc";

        String content = "<html><body><p style=\"font-family: 宋体; font-size: 16px;text-align: right;\">\n" +
                "    密级：<span style=\"text-decoration: underline;\">&nbsp; &nbsp; &nbsp;&nbsp;</span>\n" +
                "</p>\n" +
                "<p style=\"font-family: 宋体; font-size: 16px;text-align: right;\">\n" +
                "    编号：<span style=\"text-decoration: underline;\">&nbsp; &nbsp; &nbsp;&nbsp;</span>\n" +
                "</p>\n" +
                "<p>&nbsp;</p>\n" +
                "<p>&nbsp;</p>\n" +
                "<p>&nbsp;</p>\n" +
                "<p style=\"font-family: 宋体; font-size: 48px;text-align: center;\">\n" +
                "    <span class=\"item1\">帝国时代跟大色</span>\n" +
                "</p>\n" +
                "<p style=\"font-family: 宋体; font-size: 48px;text-align: center;\">\n" +
                "    <span>测试性试验大纲</span>\n" +
                "</p>\n" +
                "<p style=\"font-family: 宋体; font-size: 48px;text-align: center;\">\n" +
                "    <span class=\"item2\">鉴定试验</span>\n" +
                "</p>\n" +
                "<p>&nbsp;</p>\n" +
                "<p>&nbsp;</p>\n" +
                "<p style=\"font-family: 宋体; font-size: 21px;text-align: center;\">\n" +
                "    共<span style=\"text-decoration:underline;\">&nbsp;&nbsp;&nbsp;&nbsp;</span>页\n" +
                "</p></body></html>";
        String htmlFile = writeHtml(content,"C:\\Users\\yilanqunzhi\\Desktop\\","utf-8");
        JacobUtil.htmlToWord(htmlFile, wordFile);
    }



    public static String writeHtml(String content,String basePath){
        String uuid = UUID.randomUUID().toString();
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(basePath+uuid+".html");
            BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fileOutputStream,"utf-8"));
            bw.write(content);
            bw.close();
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return basePath+uuid+".html";
    }

    /**
     * 将html字符串写入html文件
     * @param content
     * @param basePath
     * @return
     */
    public static String writeHtml(String content,String basePath,String encode){
        String uuid = UUID.randomUUID().toString();
        try {
            String filePath=basePath+uuid+".html";
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fileOutputStream,encode));
            bw.write(content);
            bw.close();
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return basePath+uuid+".html";
    }

    /**
     * JACOB方式
     * notes:需要将jacob.dll拷贝到windows/system32或者项目所在jre\bin目录下面(比如我的Eclipse正在用的Jre路径是D:\Java\jdk1.7.0_17\jre\bin)。
     * @param htmlFile html静态页面路径
     * @param wordFile 要生成的word文档路径
     */
    public static void htmlToWord(String htmlFile, String wordFile) {
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        app.setProperty("Visible", new Variant(false));
        System.out.println("*****正在转换...*****");
        try {
            Dispatch wordDoc = app.getProperty("Documents").toDispatch();
            wordDoc = Dispatch.invoke(wordDoc, "Add", Dispatch.Method, new Object[0], new int[1]).toDispatch();
            Dispatch.invoke(app.getProperty("Selection").toDispatch(), "InsertFile", Dispatch.Method, new Object[] { htmlFile, "", new Variant(false), new Variant(false), new Variant(false) }, new int[3]);
            Dispatch.invoke(wordDoc, "SaveAs", Dispatch.Method, new Object[] {wordFile, new Variant(HTML_WORD)}, new int[1]);
            Dispatch.call(wordDoc, "Close", new Variant(false));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            app.invoke("Quit", new Variant[] {});
        }
        System.out.println("*****转换完毕********");
    }
}
