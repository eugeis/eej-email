/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.outlook;

import java.io.IOException;
import java.io.Reader;
import java.io.StringReader;
import java.io.StringWriter;
import java.io.Writer;

import javax.swing.JEditorPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.EditorKit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class RichTextToHtml {
  private final Logger logger = LoggerFactory.getLogger(RichTextToHtml.class);

  public RichTextToHtml() {
    super();
  }

  public String rtfToHtml(String richText) throws IOException {
    String ret = richText;
    Reader reader = new StringReader(richText);
    try {
      ret = rtfToHtml(reader);
    } catch (Exception e) {
      logger.error("Convertion not possible because of exception: {}", e);
    } finally {
      reader.close();
    }
    return ret;
  }

  public String rtfToHtml(Reader rtf) throws Exception, BadLocationException {
    String ret = null;
    JEditorPane p = new JEditorPane();
    p.setContentType("text/rtf");
    EditorKit kitRtf = p.getEditorKitForContentType("text/rtf");
    kitRtf.read(rtf, p.getDocument(), 0);
    kitRtf = null;
    EditorKit kitHtml = p.getEditorKitForContentType("text/html");
    Writer writer = new StringWriter();
    kitHtml.write(writer, p.getDocument(), 0, p.getDocument().getLength());
    ret = writer.toString();

    return ret;
  }
}
