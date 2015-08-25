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

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Eugen Eisler
 */
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
