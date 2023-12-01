package indi.lyl.docs.poi.word.docx;

import indi.lyl.docs.poi.word.doc.DocDemo;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.lang.annotation.ElementType;
import java.nio.charset.StandardCharsets;
import java.util.List;

public class DocxDemo {

    Writer _out;
    XWPFDocument _doc;

    public static void main(String[] args) throws IOException {
        try (InputStream is = new FileInputStream("E:\\poi-test\\1.docx");
             OutputStream out = new FileOutputStream("E:\\poi-test\\test_docx.txt")) {
            new DocxDemo(new XWPFDocument(is), out);
        }
    }

    public DocxDemo(XWPFDocument doc, OutputStream stream) throws IOException
    {
        _out = new OutputStreamWriter (stream, StandardCharsets.UTF_8);
        _doc = doc;
        List<XWPFParagraph> paragraphs =  _doc.getParagraphs();
        List<IBodyElement> bodyElementLsit = doc.getBodyElements();
        for(IBodyElement element : bodyElementLsit){

            if(element instanceof XWPFParagraph){
                // paragragh
                System.err.println("段落：");
                System.out.println(((XWPFParagraph) element).getText());
            }else if(element instanceof XWPFTable) {
                System.err.println("表格：");
                System.out.println(((XWPFTable) element).getText());
            }else if(element instanceof XWPFSDT) {
                System.err.println("目录：");
                System.out.println(((XWPFSDT) element).getContent().getText());
            }else{
                System.err.println("其他");
            }

        }
        if(true){
            return;
        }
        for(XWPFParagraph paragraph : paragraphs){
            System.out.println(paragraph.getRuns());
            String s1 = paragraph.getText();
            System.out.println("s1:" + s1);
            String s2 = paragraph.getParagraphText();
            System.out.println("s2:" + s2);
            String s3 = paragraph.getPictureText();
            System.out.println("s3:" + s3);
            _out.write(paragraph.getText());
        }
    }
}
