/*
 * Controlguide
 * Copyright (c) Siemens AG 2015, All Rights Reserved, Confidential
 */
package ee.email.core;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TextToHtml implements Converter<String, String> {

  protected Pattern linkPattern = Pattern.compile("(?i)\\b((?:https?://|www\\d{0,3}[.]|[a-z0-9.\\-]+[.][a-z]{2,4}/)(?:[^\\s()<>]+|\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\))+(?:\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\)|[^\\s`!()\\[\\]{};:\'\".,<>?«»“”‘’]))");

  @Override
  public String convert(String s) {
    StringBuilder ret = new StringBuilder();
    boolean previousWasASpace = false;
    for (char c : s.toCharArray()) {
      if (c == ' ') {
        if (previousWasASpace) {
          ret.append("&nbsp;");
          previousWasASpace = false;
          continue;
        }
        previousWasASpace = true;
      } else {
        previousWasASpace = false;
      }
      switch (c) {
      case '<':
        ret.append("&lt;");
        break;
      case '>':
        ret.append("&gt;");
        break;
      case '&':
        ret.append("&amp;");
        break;
      case '"':
        ret.append("&quot;");
        break;
      case '\n':
        ret.append("<br>");
        break;
      // We need Tab support here, because we print StackTraces as HTML
      case '\t':
        ret.append("&nbsp; &nbsp; &nbsp;");
        break;
      default:
        if (c < 128) {
          ret.append(c);
        } else {
          ret.append("&#").append((int) c).append(";");
        }
      }
    }

    String converted = ret.toString();

    Matcher matcher = linkPattern.matcher(converted);
    converted = matcher.replaceAll("<a href=\"$1\">$1</a>");
    return converted;
  }
}
