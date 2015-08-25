/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.core;

import static org.junit.Assert.*;

import org.junit.Test;

public class RichTextToHtmlTest {

  /**
   * Test method for {@link ee.email.core.RichTextToHtml#convert(java.lang.String)}.
   */
  @Test
  public void testConvert() {
    RichTextToHtml converter = new RichTextToHtml();
    String html = converter.convert("{\\rtf1\\ansi \\deflang1033\\deff0{\\fonttbl {\\f0\\froman \\fcharset0 \\fprq2 Times New Roman;}}{\\colortbl;\\red0\\green0\\blue0;} {\\stylesheet{\\fs20 \\snext0 Normal;}} {\\plain \\fs26 \\strike\\fs26 This is supposed to be strike-through.}{\\plain \\fs26 \\fs26  } {\\plain \\fs26 \\ul\\fs26 Underline text here} {\\plain \\fs26 \\fs26 .{\\u698\\'20}}");
    assertNotNull(html);
    assertFalse(html.isEmpty());
  }

}
