package com.edhn.tools.convert;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;

import com.edhn.commons.io.FileBatchProcessor;
import com.edhn.tools.office.WordWriter;


/**
 * TextConvertor
 * 
 * @author fengyqØ
 * @version 1.0
 * @date 2019-12-04
 * 
 */
public class TextConvertor extends FileBatchProcessor  {
    
    public boolean needUpdate(File f, int index) {
        File outFile = getOutputFile(f, ".docx");
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
            String content = FileUtils.readFileToString(f, getEncoding());
            File outputFile = getOutputFile(f, ".docx");
            WordWriter writer = new WordWriter();
            writer.create(outputFile.getAbsolutePath(), WordWriter.DOC_MODE_2007, content);
            System.out.println(String.format("【转化文件】第%d个文件：%s", index, f.getAbsolutePath()));
        } catch (IOException e) {
            System.out.println(String.format("【转化文件出错】第%d个文件：%s", index, f.getAbsolutePath()));
        }
    }

}
