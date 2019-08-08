package com.elens.convert.model;

/**
 * @Author: WangJian
 * @Date: 2019/8/6 14:47
 */
public class ParamVo {
    private String name;
    private String content;
    private String wordFile;
    private String htmlFile;

    public ParamVo(String name, String content, String wordFile, String htmlFile) {
        this.name = name;
        this.content = content;
        this.wordFile = wordFile;
        this.htmlFile = htmlFile;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getWordFile() {
        return wordFile;
    }

    public void setWordFile(String wordFile) {
        this.wordFile = wordFile;
    }

    public String getHtmlFile() {
        return htmlFile;
    }

    public void setHtmlFile(String htmlFile) {
        this.htmlFile = htmlFile;
    }
}
