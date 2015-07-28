package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlMailingAddress</a>
 *      </p>
 * @author eugeis
 */
public enum OlMailingAddressEnum {
  olNone(0), olHome(1), olBusiness(2), olOther(3);

  private final int value;

  private OlMailingAddressEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlMailingAddressEnum findEnum(Integer value) {

    if (value != null) {
      for (OlMailingAddressEnum objEnum : values()) {
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

  public boolean isOlNone() {

    return olNone == this;
  }

  public boolean isOlHome() {

    return olHome == this;
  }

  public boolean isOlBusiness() {

    return olBusiness == this;
  }

  public boolean isOlOther() {

    return olOther == this;
  }
}
