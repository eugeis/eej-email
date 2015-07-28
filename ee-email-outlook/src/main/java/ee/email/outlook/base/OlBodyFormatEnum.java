package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlBodyFormat</a>
 *      </p>
 * @author eugeis
 */
public enum OlBodyFormatEnum {
  olFormatUnspecified(0), olFormatPlain(1), olFormatHTML(2), olFormatRichText(3);

  private final int value;

  private OlBodyFormatEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlBodyFormatEnum findEnum(Integer value) {

    if (value != null) {
      for (OlBodyFormatEnum objEnum : values()) {
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

  public boolean isOlFormatUnspecified() {

    return olFormatUnspecified == this;
  }

  public boolean isOlFormatPlain() {

    return olFormatPlain == this;
  }

  public boolean isOlFormatHTML() {

    return olFormatHTML == this;
  }

  public boolean isOlFormatRichText() {

    return olFormatRichText == this;
  }
}
