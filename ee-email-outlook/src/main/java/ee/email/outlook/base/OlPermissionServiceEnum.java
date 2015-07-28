package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlPermissionService</a>
 *      </p>
 * @author eugeis
 */
public enum OlPermissionServiceEnum {
  olUnknown(0), olWindows(1), olPassport(2);

  private final int value;

  private OlPermissionServiceEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlPermissionServiceEnum findEnum(Integer value) {

    if (value != null) {
      for (OlPermissionServiceEnum objEnum : values()) {
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

  public boolean isOlUnknown() {

    return olUnknown == this;
  }

  public boolean isOlWindows() {

    return olWindows == this;
  }

  public boolean isOlPassport() {

    return olPassport == this;
  }
}
