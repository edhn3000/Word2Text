package com.edhn.tools.writ;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.Arrays;
import java.util.HashSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

import com.edhn.commons.io.FileBatchProcessor;
import com.edhn.tools.office.WordReader;
import com.edhn.tools.prop.PropReader;
import com.google.common.io.Files;

/**
 * WordProcess
 * @author fengyq
 * @version 1.0
 *
 */
public class WritFormatName extends FileBatchProcessor {
    
    private String outEncoding = "GBK";
    static Pattern ahPattern;

    @Override
    public void processFile(File f, int index) {
        try {
            String fileName = f.getCanonicalPath();
            String text = "";
            if (fileName.toLowerCase().endsWith(".txt")) {
                text = FileUtils.readFileToString(new File(fileName), Charset.forName("GBK"));
            } else {
                WordReader reader = new WordReader();
                try {
                    reader.load(fileName);
                    text = reader.getText().trim();
                } catch (OfficeXmlFileException e) {
                    // doc扩展名但实际是docx的情况
                    if (fileName.toLowerCase().endsWith(".doc")) {
                        reader.load(fileName, WordReader.DOC_MODE_2007);
//                        System.out.println(String.format("扩展名错误，文件%s的扩展名应为.docx", fileName));
                        text = reader.getText().trim();
                    }
                } catch (InvalidOperationException e) {
                    // docx扩展名但实际是doc的情况
                    if (fileName.toLowerCase().endsWith(".docx")) {
                        reader.load(fileName, WordReader.DOC_MODE_2003);
//                        System.out.println(String.format("扩展名错误，文件%s的扩展名应为.doc", fileName));
                        text = reader.getText().trim();
                    }
                }
                reader = null;
            }
            text = text.replaceAll("[\u0020\t\u3000]+", "");
            text = text.replaceAll("\r?\n?(\r?\n)+", "\n");
            
            Matcher m = ahPattern.matcher(text);
            String ah = "";
            if (m.find()) {
                ah = m.group(1);
            } else if (f.getParentFile().getName().matches(".*\\d+号")){
                ah = f.getParentFile().getName();
            } else {
                System.out.println(String.format("无法分析案号，文件=%s", fileName));
                return;
            }
            
            if (ah != null && !"".equals(ah) && !f.getName().toLowerCase().contains(ah.toLowerCase())) {
                String newFileName = String.format("%s" + File.separator + "%s_%s.%s", 
                    f.getParent(), ah, Files.getNameWithoutExtension(f.getName()), Files.getFileExtension(f.getName()));
                System.out.println(String.format("改名，%s->%s", fileName, new File(newFileName).getName()));
                FileUtils.moveFile(f, new File(newFileName));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    
    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        if (args.length == 0) {
            return ;
        }
        
        PropReader reader = new PropReader();
        reader.loadPropResource("/prop/writRegex.ini", false, "GBK");
//        reader.loadPropFile("e:\\MyProgram\\regex\\writRegex.ini", false, "GBK");
        String regex = reader.getPropValue("案号");
        if (regex == null) {
            System.out.println("无法加载规则文件/prop/writRegex.ini");
            return ;
        }
        ahPattern = Pattern.compile(regex);
        
        String path = "";
        WritFormatName convertor = new WritFormatName();
        
        for (int i = 0; i < args.length; i++) {
            String arg = args[i];
            if (arg.startsWith("-")) {
                if (arg.startsWith("-encoding=")) {
                    convertor.setOutEncoding(arg.substring("-encoding=".length()).trim());
                } else if (arg.toLowerCase().startsWith("-update")) {
                    convertor.setUpdate(true);
                }
            } else {
                path = arg;
                if (path.endsWith("\\") || path.endsWith("/")) {
                	path = path.substring(0, path.length() - 1);
                }
            }
        }
        
        File dir = new File(path);
        if (!dir.exists()) {
            System.out.println("传入的路径不存在！" + path);
            return ;
        }
        if (!dir.isDirectory()) {
            System.out.println("传入的路径不是目录！" + path);
            return ;
        }

        convertor.processDir(path, new HashSet<String>(Arrays.asList(new String[]{"docx", "doc", "txt"})));
        
        System.out.println("执行完毕!");
    }

    /**
     * @return the outEncoding
     */
    public String getOutEncoding() {
        return outEncoding;
    }

    /**
     * @param outEncoding the outEncoding to set
     */
    public void setOutEncoding(String outEncoding) {
        this.outEncoding = outEncoding;
    }
    
}
