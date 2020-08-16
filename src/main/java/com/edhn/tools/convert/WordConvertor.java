package com.edhn.tools.convert;

import java.io.File;
import java.nio.charset.Charset;

import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

import com.edhn.commons.io.FileBatchProcessor;
import com.edhn.tools.office.WordReader;

/**
 * WordConvertor
 * @author fengyq
 * @version 1.0
 *
 */
public class WordConvertor extends FileBatchProcessor {
    
    @Override
    public boolean needUpdate(File f, int index) {
        File outFile = getOutputFile(f, ".txt");
        if (outFile.exists()) {
            if (outFile.lastModified() > f.lastModified()) {
                System.out.println(String.format("update模式，跳过第%d个文件：%s", index, f.getAbsolutePath()));
                return false;
            }
        }
        return true;
    }
    
    @Override
    public void processFile(File f, int index) {
        try {
            String fileName = f.getCanonicalPath();
            File outFile = getOutputFile(f, ".txt");

            WordReader reader = new WordReader();
            String text = "";
            String extHint = "";
            try {
                reader.load(fileName);
                text = reader.getText().trim();
                FileUtils.write(outFile, text, getEncoding());
                System.out.println(String.format("【转化文件】第%d个文件：%s", index, fileName));
            } catch (OfficeXmlFileException e) {
                // doc扩展名但实际是docx的情况
                if (fileName.toLowerCase().endsWith(".doc")) {
                    reader.load(fileName, WordReader.DOC_MODE_2007);
                    extHint = String.format("扩展名错误，该文件扩展名应为.docx", fileName);
                    text = reader.getText().trim();
                    FileUtils.write(outFile, text, Charset.forName("GBK"));
                    System.out.println(String.format("【转化文件出错】第%d个文件：%s（%s）", index, fileName, extHint));
                }
            } catch (InvalidOperationException e) {
                // docx扩展名但实际是doc的情况
                if (fileName.toLowerCase().endsWith(".docx")) {
                    reader.load(fileName, WordReader.DOC_MODE_2003);
                    extHint = String.format("扩展名错误，该文件扩展名应为.doc", fileName);
                    text = reader.getText().trim();
                    FileUtils.write(outFile, text, Charset.forName("GBK"));
                    System.out.println(String.format("【转化文件出错】第%d个文件：%s（%s）", index, fileName, extHint));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
