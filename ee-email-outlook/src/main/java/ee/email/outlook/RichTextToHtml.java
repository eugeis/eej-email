/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.outlook;

import java.io.StringWriter;
import java.io.Writer;

import javax.swing.JEditorPane;
import javax.swing.text.EditorKit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class RichTextToHtml {
  private final Logger logger = LoggerFactory.getLogger(RichTextToHtml.class);

  protected JEditorPane p;

  public RichTextToHtml() {
    super();
    p = new JEditorPane();
    p.setContentType("text/rtf");
  }

  public String rtfToHtml(String richText) {
    String ret = richText;

    try {
      p.setText(richText);
      EditorKit kitHtml = p.getEditorKitForContentType("text/html");
      Writer writer = new StringWriter();
      kitHtml.write(writer, p.getDocument(), 0, p.getDocument().getLength());
      ret = writer.toString();
    } catch (Exception e) {
      logger.error("Convertion not possible because of exception: {}", e);
    } finally {
      p.setText("");
    }
    return ret;
  }
}
