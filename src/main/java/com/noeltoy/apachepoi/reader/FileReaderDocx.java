package com.noeltoy.apachepoi.reader;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class FileReaderDocx {
    public void docxReader(){
        String filePath = "/home/noel/Downloads/file-sample_100kB.docx";
        try {
            XWPFDocument docx = new XWPFDocument(new FileInputStream(filePath));
            docx = replaceText(docx, "$VAR", "MyValue1");
            saveWord(filePath, docx);
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }
    private static XWPFDocument replaceText(XWPFDocument docx, String findText, String replaceText){
        for(XWPFParagraph para : docx.getParagraphs()){
            for (XWPFRun run:para.getRuns()){
                String text = run.getText(0);
                if(text!=null && text.contains(findText)) {
                    text = text.replace(findText,replaceText);
                    run.setText(text,0);
                }
            }
        }
        return docx;
    }

    private static void saveWord(String filePath, XWPFDocument docx) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            docx.write(out);
        }
        finally{
            assert out != null;
            out.close();
        }
    }
}
