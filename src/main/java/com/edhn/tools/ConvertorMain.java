package com.edhn.tools;

import java.io.File;
import java.nio.charset.Charset;
import java.util.Arrays;
import java.util.HashSet;

import com.edhn.commons.io.FileBatchProcessor;
import com.edhn.tools.convert.TextConvertor;
import com.edhn.tools.convert.WordConvertor;

public class ConvertorMain {   
    
    public static void main(String[] args) {
        if (args.length == 0) {
            System.out.println("使用方法：Word2Text path [-encoding=] [-update] [-outdir=]");
            System.out.println("    path        源文件所在路径，注意结尾不要带\\");
            System.out.println("    -encoding   转换后输出文件的字符集，如GBK、UTF-8等");
            System.out.println("    -update     已转换过且word文档未被更新的时不重复转换");
            System.out.println("    -outdir     输出路径，默认是跟原文件同目录");
            System.out.println("    -txt2docx   反转换，文本转word文档docx格式");
            return ;
        }
        String path = "";
        
        String encoding = null;
        String outDir = "";
        boolean update = false;
        boolean txt2Docx = false;
        
        for (int i = 0; i < args.length; i++) {
            String arg = args[i];
            if (arg.startsWith("-")) {
                if (arg.toLowerCase().startsWith("-encoding=")) {
                    encoding = arg.substring("-encoding=".length()).trim();
                } else if (arg.toLowerCase().startsWith("-update")) {
                    update = true;
                } else if (arg.toLowerCase().startsWith("-outdir=")) {
                    outDir = arg.substring("-outdir=".length()).trim();
                } else if (arg.toLowerCase().startsWith("-txt2docx")) {
                    txt2Docx = true;
                }
            } else {
                path = arg;
                if (path.endsWith(File.separator))
                    path = path.substring(0, path.length() - 1);
            }
        }
        
        File f = new File(path);
        if (!f.exists()) {
            System.out.println("传入的路径不存在！" + path);
            return ;
        }

        FileBatchProcessor convertor = null;
        if (txt2Docx) {
            convertor = new TextConvertor();
        } else {
            convertor = new WordConvertor();
        }
        convertor.setOutputRoot(outDir);
        convertor.setUpdate(update);
        if (encoding != null) {
            convertor.setEncoding(Charset.forName(encoding));
        }
        convertor.setInputRoot(path);
        
        if (f.isDirectory()) {
	        if (txt2Docx) {
	            convertor.processDir(path, new HashSet<String>(Arrays.asList(new String[]{"txt"})));
	        } else {
	            convertor.processDir(path, new HashSet<String>(Arrays.asList(new String[]{"docx", "doc"})));
	        }
	    } else {
	    	convertor.processFile(f, 0);
	    }
        
        
        System.out.println("执行完毕!");
    }

}
