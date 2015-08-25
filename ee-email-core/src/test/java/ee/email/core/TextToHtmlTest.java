/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.core;

import static org.junit.Assert.*;

import org.junit.Test;

public class TextToHtmlTest {

  /**
   * Test method for {@link ee.email.core.RichTextToHtml#convert(java.lang.String)}.
   */
  @Test
  public void testConvert() {
    TextToHtml converter = new TextToHtml();
    String html = converter.convert("Hello\n, here is a test.");
    assertNotNull(html);
    assertFalse(html.isEmpty());
  }
}
