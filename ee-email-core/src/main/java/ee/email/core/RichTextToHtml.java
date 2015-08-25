/*
 * Copyright 2015-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package ee.email.core;

import java.io.IOException;
import java.io.Reader;
import java.io.StringReader;
import java.io.StringWriter;

import javax.swing.JEditorPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.EditorKit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Eugen Eisler
 */
public class RichTextToHtml implements Converter<String, String> {
  private final Logger logger = LoggerFactory.getLogger(RichTextToHtml.class);

  public RichTextToHtml() {
    super();
  }

  @Override
  public String convert(String richText) {
    String ret = richText;
    Reader reader = new StringReader(richText);
    try {
      ret = rtfToHtml(reader);
    } catch (Exception e) {
      logger.error("Convertion not possible because of exception: {}", e);
    } finally {
      try {
        reader.close();
      } catch (IOException e) {
      }
    }
    return ret;
  }

  protected String rtfToHtml(Reader rtf) throws Exception, BadLocationException {
    String ret = null;
    JEditorPane p = new JEditorPane();
    p.setContentType("text/rtf");
    EditorKit kitRtf = p.getEditorKitForContentType("text/rtf");
    kitRtf.read(rtf, p.getDocument(), 0);
    kitRtf = null;
    EditorKit kitHtml = p.getEditorKitForContentType("text/html");
    StringWriter writer = new StringWriter();
    kitHtml.write(writer, p.getDocument(), 0, p.getDocument().getLength());
    ret = writer.toString();

    return ret;
  }
}
