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
package ee.email.model;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlBodyFormat</a>
 *      </p>
 * @author eugeis
 */
public enum BodyFormatEnum {
  Unspecified(0), Plain(1), HTML(2), RichText(3);

  private final int value;

  private BodyFormatEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static BodyFormatEnum findEnum(Integer value) {

    if (value != null) {
      for (BodyFormatEnum objEnum : values()) {
        if (objEnum.value == value) {
          return objEnum;
        }
      }
    }
    return null;
  }

  public boolean isValue(int value) {

    return this.value == value;
  }

  public boolean isUnspecified() {

    return Unspecified == this;
  }

  public boolean isPlain() {

    return Plain == this;
  }

  public boolean isHTML() {

    return HTML == this;
  }

  public boolean isRichText() {

    return RichText == this;
  }
}
