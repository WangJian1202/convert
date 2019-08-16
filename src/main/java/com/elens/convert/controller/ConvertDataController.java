package com.elens.convert.controller;


import com.elens.convert.utils.CompressedFileUtil;
import com.elens.convert.utils.EncodingDetect;
import com.elens.convert.utils.JacobUtil;
import com.elens.convert.utils.Office2HtmlUtil;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: WangJian
 * @Date: 2019/8/6 10:13
 */
@RestController
@RequestMapping("/convert")
public class ConvertDataController {


    @RequestMapping(value = "/exportWord", method = RequestMethod.POST)
    @ResponseBody
    public void exportWord(MultipartFile uploadFile, HttpServletRequest request, HttpServletResponse response) {
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        File filePath = null;
        File saveFile = null;
        String uploadPath = System.getProperty("user.dir") + "\\upload\\";
        try {
            String name = uploadFile.getOriginalFilename();
            saveFile = new File(uploadPath, name);
            FileUtils.copyInputStreamToFile(uploadFile.getInputStream(), saveFile);
            String javaEncode = EncodingDetect.getJavaEncode(saveFile.getAbsolutePath());
            System.out.println(javaEncode);
            String content = getString(saveFile, javaEncode);
            saveFile.delete();
            System.out.println(name);
            System.out.println(content);
            if (!content.equals("")) {
                if (!content.contains("<body>")) {
                    content = "<body>" + content + "</body>";
                }
                if (!content.contains("</html>")) {
                    content = "<html>" + content + "</html>";
                }
                String downloadPath = System.getProperty("user.dir") + "\\" + System.currentTimeMillis() + "\\";
                filePath = new File(downloadPath);
                if (!filePath.exists()) {
                    filePath.mkdirs();
                }
                String htmlFile = JacobUtil.writeHtml(content, downloadPath, javaEncode);
                System.out.println(htmlFile);
                name = name.split("\\.")[0];
                String wordFile = downloadPath + name + ".docx";
                JacobUtil.htmlToWord(htmlFile, wordFile);
                File file = new File(wordFile);
                response.setContentType("application/x-msdownload;");
                response.setHeader("Content-disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(name + ".docx", "UTF-8"));
                response.setHeader("Content-Length", String.valueOf(file.length()));
                bis = new BufferedInputStream(new FileInputStream(file));
                bos = new BufferedOutputStream(response.getOutputStream());
                byte[] buff = new byte[2048];
                int bytesRead;
                while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                    bos.write(buff, 0, bytesRead);
                }
                bos.flush();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (bis != null) {
                    bis.close();
                }
                if (bos != null) {
                    bos.close();
                }
                if (filePath != null) {
                    filePath.delete();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    //    从文件获取字符串
    private String getString(File saveFile, String encode) throws IOException {
        StringBuffer buffer = new StringBuffer();
        FileInputStream in = new FileInputStream(saveFile);
        BufferedReader bf = new BufferedReader(new InputStreamReader(in, encode));
        String s = null;
        while ((s = bf.readLine()) != null) {
            buffer.append(s.trim());
        }
        bf.close();
        return buffer.toString();
    }


    @RequestMapping(value = "/exportHtml", method = RequestMethod.POST)
    @ResponseBody
    public void exportHtml(MultipartFile file, HttpServletResponse response) {
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        File targetFile = null;
        File saveFile = null;
        String projectPath = System.getProperty("user.dir");
//            上传文件保存路径
        String path = projectPath + "\\upload\\";
//            生成html文件及其图片文件
        String createdPath = projectPath + "\\created\\";
//            生成的打包文件保存路径
        String targetPath = projectPath + "\\zip\\";
        try {
            System.out.println("path:" + path);
            String name = file.getOriginalFilename();
            System.out.println("name:" + name);
//          将上传文件保存到upload路径下
            saveFile = new File(path, name);
            FileUtils.copyInputStreamToFile(file.getInputStream(), saveFile);
            name = name.split("\\.")[0];
            String htmlFile = createdPath + name + ".html";
            System.out.println("htmlFile:" + htmlFile);
//            将上传文件转成html并将生成的html及图片文件保存至created路径下
            Office2HtmlUtil.wordToHtml(saveFile.getPath(), htmlFile);
//            打包created文件夹下所有文件
            String zipName = name + ".zip";
            CompressedFileUtil.compressedFile(createdPath, targetPath, zipName);
//            返回下载文件
            targetFile = new File(targetPath, zipName);
            response.setContentType("application/x-msdownload;");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(zipName, "UTF-8"));
            response.setHeader("Content-Length", String.valueOf(targetFile.length()));
            bis = new BufferedInputStream(new FileInputStream(targetFile));
            bos = new BufferedOutputStream(response.getOutputStream());
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
            bos.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (bis != null) {
                    bis.close();
                }
                if (bos != null) {
                    bos.close();
                }
//          删除 zip upload created包下的所有文件
                delZSPic(path);
                delZSPic(targetPath);
                delZSPic(createdPath);
            } catch (Exception e) {
                e.printStackTrace();
            }


        }
    }

    @RequestMapping(value = "/exportHtml2", method = RequestMethod.POST)
    @ResponseBody
    public void exportHtml2(MultipartFile file, HttpServletResponse response) {
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        File targetFile = null;
        File saveFile = null;
        String projectPath = System.getProperty("user.dir");
//            上传文件保存路径
        String path = projectPath + "\\upload\\";
//            生成html文件及其图片文件
        String createdPath = projectPath + "\\created\\";
        try {
            String name = file.getOriginalFilename();
            System.out.println("name:" + name);
//          将上传文件保存到upload路径下
            saveFile = new File(path, name);
            FileUtils.copyInputStreamToFile(file.getInputStream(), saveFile);
            name = name.split("\\.")[0];
            String htmlFile = createdPath + name + ".html";
            System.out.println("htmlFile:" + htmlFile);
//            将上传文件转成html并将生成的html及图片文件保存至created路径下
//            Office2HtmlUtil.wordToHtml(saveFile.getPath(), htmlFile);
            Boolean aBoolean = word2htmlByPython(htmlFile,saveFile.getAbsolutePath());
            System.out.println(aBoolean);
            targetFile = new File(htmlFile);
            if(aBoolean){
                String outputName=name+".html";
                response.setContentType("application/x-msdownload;");
                response.setHeader("Content-disposition", "attachment;filename*=utf-8'zh_cn'" + URLEncoder.encode(outputName, "UTF-8"));
                response.setHeader("Content-Length", String.valueOf(targetFile.length()));
                bis = new BufferedInputStream(new FileInputStream(targetFile));
                bos = new BufferedOutputStream(response.getOutputStream());
                byte[] buff = new byte[2048];
                int bytesRead;
                while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                    bos.write(buff, 0, bytesRead);
                }
                bos.flush();
            }
            delZSPic(path);
            delZSPic(createdPath);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (bis != null) {
                    bis.close();
                }
                if (bos != null) {
                    bos.close();
                }

            } catch (Exception e) {
                e.printStackTrace();
            }


        }


    }


    public Boolean word2htmlByPython(String inputPath,String outputPath) {
        String property = System.getProperty("user.dir");
        String pythonFile = property + "\\python\\" + "html2word.py";
        String[] commands = new String[]{"python", pythonFile, inputPath, outputPath};
        System.out.println(StringUtils.join(commands, " "));
        boolean flag = false;
        Process proc = null;
        try {
            proc = Runtime.getRuntime().exec(commands);
            //用输入输出流来截取结果
            BufferedReader in = new BufferedReader(new InputStreamReader(proc.getInputStream(), "utf-8"));
            String line = null;
            String msg = "";
            while ((line = in.readLine()) != null) {
                System.out.println(line);
                msg = line;
            }
            if (msg.equals("data end")) {
                flag = true;
            }
            in.close();
            proc.waitFor();
            proc.destroy();
        } catch (Exception e) {
            e.printStackTrace();
            if (proc != null) {
                proc.destroy();
            }
        }
        return flag;
    }








    //    删除某个路径下所有文件
    private boolean delZSPic(String filePath) {
        boolean flag = true;
        if (filePath != null) {
            File file = new File(filePath);
            if (file.exists()) {
                File[] filePaths = file.listFiles();
                for (File f : filePaths) {
                    if (f.isFile()) {
                        f.delete();
                    }
                    if (f.isDirectory()) {
                        String fpath = f.getPath();
                        delZSPic(fpath);
                        f.delete();
                    }
                }
            }
        } else {
            flag = false;
        }
        return flag;
    }


}
