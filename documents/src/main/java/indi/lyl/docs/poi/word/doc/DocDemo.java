package indi.lyl.docs.poi.word.doc;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class DocDemo {
    Writer _out;
    HWPFDocument _doc;

    public DocDemo(HWPFDocument doc, OutputStream stream) throws IOException
    {
        _out = new OutputStreamWriter (stream, StandardCharsets.UTF_8);
        _doc = doc;

//        init ();
//        openDocument ();
//        openBody ();

        Range r = doc.getRange ();
        writePlainText(r.text());
        if(true){
            _out.flush();
            return;
        }
        System.out.println("text blow is from range object:");
        System.out.println(r.text());
        StyleSheet styleSheet = doc.getStyleSheet ();

        int sectionLevel = 0;
        int lenParagraph = r.numParagraphs ();
        boolean inCode = false;
        for (int x = 0; x < lenParagraph; x++)
        {
            Paragraph p = r.getParagraph (x);
            String text = p.text ();
            if (text.trim ().length () == 0){
                continue;
            } else {
                CharacterRun run = p.getCharacterRun(0);
                System.out.println(run.text());
                writePlainText(x + p.text());
                if(true){
                    continue;
                }
            }

            StyleDescription paragraphStyle = styleSheet.getStyleDescription (p.
                    getStyleIndex ());
            String styleName = paragraphStyle.getName();
            if (styleName.startsWith ("Heading")){
                if (inCode)
                {
                    closeSource();
                    inCode = false;
                }

                int headerLevel = Integer.parseInt (styleName.substring (8));
                if (headerLevel > sectionLevel)
                {
                    openSection ();
                }
                else
                {
                    for (int y = 0; y < (sectionLevel - headerLevel) + 1; y++)
                    {
                        closeSection ();
                    }
                    openSection ();
                }
                sectionLevel = headerLevel;
                openTitle ();
                writePlainText (text);
                closeTitle ();
            } else {
                int cruns = p.numCharacterRuns ();
                CharacterRun run = p.getCharacterRun (0);
                String fontName = run.getFontName();
                if (fontName.startsWith ("Courier"))
                {
                    if (!inCode)
                    {
                        openSource ();
                        inCode = true;
                    }
                    writePlainText (p.text());
                }
                else
                {
                    if (inCode)
                    {
                        inCode = false;
                        closeSource();
                    }
                    openParagraph();
                    writePlainText(p.text());
                    closeParagraph();
                }
            }
        }
//        for (int x = 0; x < sectionLevel; x++)
//        {
//            closeSection();
//        }
//        closeBody();
//        closeDocument();
        _out.flush();

    }

    public void init ()
            throws IOException
    {
        _out.write ("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n");
        _out.write ("<!DOCTYPE document PUBLIC \"-//APACHE//DTD Documentation V1.1//EN\" \"./dtd/document-v11.dtd\">\r\n");
    }

    public void openDocument ()
            throws IOException
    {
        _out.write ("<document>\r\n");
    }
    public void closeDocument ()
            throws IOException
    {
        _out.write ("</document>\r\n");
    }


    public void openBody ()
            throws IOException
    {
        _out.write ("<body>\r\n");
    }

    public void closeBody ()
            throws IOException
    {
        _out.write ("</body>\r\n");
    }


    public void openSection ()
            throws IOException
    {
        _out.write ("<section>");

    }

    public void closeSection ()
            throws IOException
    {
        _out.write ("</section>");

    }

    public void openTitle ()
            throws IOException
    {
        _out.write ("<title>");
    }

    public void closeTitle ()
            throws IOException
    {
        _out.write ("</title>");
    }

    public void writePlainText (String text)
            throws IOException
    {
        _out.write (text);
    }

    public void openParagraph ()
            throws IOException
    {
        _out.write ("<p>");
    }

    public void closeParagraph ()
            throws IOException
    {
        _out.write ("</p>");
    }

    public void openSource ()
            throws IOException
    {
        _out.write ("<source><![CDATA[");
    }
    public void closeSource ()
            throws IOException
    {
        _out.write ("]]></source>");
    }


    public static void main(String[] args) throws IOException {
        try (InputStream is = new FileInputStream("E:\\poi-test\\1.doc");
             OutputStream out = new FileOutputStream("E:\\poi-test\\test_doc.txt")) {
            new DocDemo(new HWPFDocument(is), out);
        }
    }
}
