package com.elens.convert.doc;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * @Author: WangJian
 * @Date: 2019/8/20 13:41
 */

public class XwpfUtils {

    public static void main(String[] args) throws Exception {
        String title="人物01,日期,地址,人物02,人物03";
        ArrayList<LinkedHashMap<Object, Object>> resultByTitle = getResultByTitle(title,"C:\\Users\\yilanqunzhi\\Desktop\\test.docx");
        writeResult(resultByTitle,title);
    }
//  生成结果文件
    public static void writeResult(List<LinkedHashMap<Object, Object>> maps,String title)throws Exception{
        String[] split = title.split(",");
        int rows = maps.size();
        XWPFDocument doc = new XWPFDocument();
        XWPFTable table = doc.createTable(rows, split.length-1);
        List<XWPFTableRow> rowList = table.getRows();
        CTTblPr tablePr  = table.getCTTbl().addNewTblPr();
        CTTblWidth width  = tablePr.addNewTblW();
        width.setW(BigInteger.valueOf(8000));
        XWPFTableRow row;
        List<XWPFTableCell> cells;
        XWPFTableCell cell;
        int rowSize = rowList.size();
        int cellSize;
        for (int i = 0; i <rowSize ; i++) {
            row=rowList.get(i);
//            新增单元格
            row.addNewTableCell();
            row.setHeight(500);
            //行属性
//       CTTrPr rowPr = row.getCtRow().addNewTrPr();
            //这种方式是可以获取到新增的cell的。
//       List<CTTc> list = row.getCtRow().getTcList();
            cells = row.getTableCells();
            cellSize=cells.size();
            for (int j = 0; j <cellSize ; j++) {
                cell=cells.get(j);
                if((i+j)%2==0){
                    cell.setColor("66ff66");
                }else{
                    cell.setColor("66ff66");
                }
                //单元格属性
                CTTcPr cellPr = cell.getCTTc().addNewTcPr();
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//                if (j == 3) {
//                    //设置宽度
//                    cellPr.addNewTcW().setW(BigInteger.valueOf(1000));
//                }
                cellPr.addNewTcW().setW(BigInteger.valueOf(1500));
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                if(i==0){
                    cell.setText(split[j]);
                }else{
                    LinkedHashMap<Object, Object> map = maps.get(i);
                    try{
                        String value = map.get(split[j]).toString();
                        cell.setText(value);
                    }catch (Exception e){
                        cell.setText("");
                    }
                }
            }
        }
        OutputStream os = new FileOutputStream("C:\\Users\\yilanqunzhi\\Desktop\\test_result.docx");
        doc.write(os);
        os.close();
    }
//    根据title重word中抽取带下划线信息
    public static  ArrayList<LinkedHashMap<Object, Object>> getResultByTitle(String title,String filePath){
        String s="";
        try {
            String[] split = title.split(",");
            InputStream is = new FileInputStream(filePath);
            XWPFDocument doc = new XWPFDocument(is);
            List<XWPFParagraph> paras  = doc.getParagraphs();
//        ArrayList<List> list = new ArrayList<>();
            ArrayList<LinkedHashMap<Object, Object>> maps = new ArrayList<>();
            StringBuffer stringBuffer = new StringBuffer();
            for (XWPFParagraph para: paras) {
                List<XWPFRun> runsLists  = para.getRuns();
                ArrayList<String> contentlist = new ArrayList<>();
//                循环标识符
                int index=0;
                for (XWPFRun runsList : runsLists) {
                    index++;
//                String c = runsList.getColor();//获取句的字体颜色
//                int f = runsList.getFontSize();//获取句中字的大小
                    s = runsList.getText(0);//获取文本内容
                    UnderlinePatterns underline = runsList.getUnderline();
                    int value = underline.getValue();
//                    将带下划线的词拼接并写入list
                    if(value!=18){
                        stringBuffer.append(s);
                    }else{
                        if(stringBuffer.length()>0){
                            contentlist.add(stringBuffer.toString());
                            stringBuffer.setLength(0);
                        }
                    }

//                  处理最后一个词带下划线的情况
                    if(index==runsLists.size()){
                        if(stringBuffer.length()>0){
                            contentlist.add(stringBuffer.toString());
                            stringBuffer.setLength(0);
                        }
                    }
                }

                if(contentlist.size()!=0){
//                list.add(contentlist);
                    LinkedHashMap<Object, Object> map = new LinkedHashMap<>();
                    for (int i = 0; i <contentlist.size() ; i++) {
//                      只取与所给字段数量相同的下划线
                        if(i>split.length-1){
                            continue;
                        }
                        map.put(split[i],contentlist.get(i));
                    }
                    maps.add(map);
                }
            }
//        System.out.println(list);
            System.out.println(maps);

//      读取表格
//        List<XWPFTable> tables = doc.getTables();
//        List<XWPFTableRow> rows;
//        List<XWPFTableCell> cells;
//        for(XWPFTable table : tables){
//            rows = table.getRows();
//            for(XWPFTableRow row : rows){
//                cells = row.getTableCells();
//                for(XWPFTableCell cell : cells){
//                    System.out.println(cell.getText());
//                }
//            }
//        }
            is.close();
            return maps;
        }catch (Exception e){
            e.printStackTrace();
            System.out.println(s);
        }
            return null;
    }

}
