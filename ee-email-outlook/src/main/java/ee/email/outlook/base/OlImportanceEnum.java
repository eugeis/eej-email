package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlImportance</a>
 *      </p>
 * @author eugeis
 */
public enum OlImportanceEnum {
  olImportanceLow(0), olImportanceNormal(1), olImportanceHigh(2);

  private final int value;

  private OlImportanceEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlImportanceEnum findEnum(Integer value) {

    if (value != null) {
      for (OlImportanceEnum objEnum : values()) {
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

  public boolean isOlImportanceLow() {

    return olImportanceLow == this;
  }

  public boolean isOlImportanceNormal() {

    return olImportanceNormal == this;
  }

  public boolean isOlImportanceHigh() {

    return olImportanceHigh == this;
  }
}
