package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlPermission</a>
 *      </p>
 * @author eugeis
 */
public enum OlPermissionEnum {
  olUnrestricted(0), olDoNotForward(1), olPermissionTemplate(2);

  private final int value;

  private OlPermissionEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlPermissionEnum findEnum(Integer value) {

    if (value != null) {
      for (OlPermissionEnum objEnum : values()) {
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

  public boolean isOlUnrestricted() {

    return olUnrestricted == this;
  }

  public boolean isOlDoNotForward() {

    return olDoNotForward == this;
  }

  public boolean isOlPermissionTemplate() {

    return olPermissionTemplate == this;
  }
}
