package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlStoreType</a>
 *      </p>
 * @author eugeis
 */
public enum OlStoreTypeEnum {
  olStoreDefault(1), olStoreUnicode(2), olStoreANSI(3);

  private final int value;

  private OlStoreTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlStoreTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlStoreTypeEnum objEnum : values()) {
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

  public boolean isOlStoreDefault() {

    return olStoreDefault == this;
  }

  public boolean isOlStoreUnicode() {

    return olStoreUnicode == this;
  }

  public boolean isOlStoreANSI() {

    return olStoreANSI == this;
  }
}
