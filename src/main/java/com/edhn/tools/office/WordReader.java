package com.edhn.tools.office;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.POITextExtractor;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

/**
 * WordReader
 * 使用POI读Word内容
 * @author fengyq
 * @version 1.0
 *
 */
public class WordReader {
    
    public final static int DOC_MODE_2003 = 1;
    public final static int DOC_MODE_2007 = 2;
    
    private String fileName;
    private POITextExtractor extractor;

    public WordReader () {
        
    }

    /**
     * @return the fileName
     */
    public String getFileName() {
        return fileName;
    }
    
    /**
     * @param fileName
     * @throws Exception
     */
    public void load(String fileName) throws Exception {
        this.fileName = fileName;
        if (fileName.toLowerCase().endsWith(".docx")) {
            this.load(fileName, DOC_MODE_2007);
        } else {
            this.load(fileName, DOC_MODE_2003);
        }
    }
    
    /**
     * @param fileName
     * @throws Exception
     */
    public void load(String fileName, int mode) throws Exception  {
        this.fileName = fileName;
        if (DOC_MODE_2007 == mode) {
            OPCPackage pkg = POIXMLDocument.openPackage(fileName);
            try{
                extractor = new XWPFWordExtractor(pkg);
            } finally {
                pkg.close();
            }
        } else if (DOC_MODE_2003 == mode){
            InputStream is = new FileInputStream(fileName);
            try {
                extractor = new WordExtractor(is);
            } finally {
                is.close();
            }
        }
    }
    
    public String getText() {
        return extractor.getText();
    }
    
    public static void main(String[] args) {
        try {
            WordReader reader = new WordReader();
            reader.load("f:\\Thunisoft\\CoCall 2.0\\ReceiveFile\\樊玲\\刑罚变更（旧）\\115009刑罚变更文书，案件由来部分应写明原一审裁判信息。其标准表述为：“×××人民法院于××××年××月××日作出了（××××）×初字第××号刑事判决，以被告人×××犯××罪，判处………（写明主刑的刑种、刑期和附加剥夺政治权利及其刑期）。”.docx");
            System.out.println(reader.getText());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
     
}
