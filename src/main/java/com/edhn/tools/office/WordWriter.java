package com.edhn.tools.office;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordWriter {
    
    public final static int DOC_MODE_2003 = 1;
    public final static int DOC_MODE_2007 = 2;
    
    POIXMLDocument docx;
    private File docFile;
    
    public WordWriter() {
        
    }
    
    public void create(String fileName, int mode, String content) throws IOException {
        docFile = new File(fileName);
        if (DOC_MODE_2003 == mode) {
//            HWPFDocument
            throw new UnsupportedEncodingException("unsupport to create doc format file!");
        } else {
            // 仅在无扩展名时添加扩展名
            if (fileName.indexOf(".") == -1) {
                fileName = fileName + ".docx";
            }
            XWPFDocument docx = new XWPFDocument();
            this.docx = docx;
            String[] lines = content.split("\r?\n");
            for (String s : lines) {
                XWPFParagraph para = docx.createParagraph();
                XWPFRun r = para.createRun();   
                r.setText(s);  
            } 
            FileOutputStream fos = new FileOutputStream(docFile);   
            docx.write(fos);   
            fos.close();   
        }
    }

}
