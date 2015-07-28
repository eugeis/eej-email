package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlWindowState</a>
 *      </p>
 * @author eugeis
 */
public enum OlWindowStateEnum {
  olMaximized(0), olMinimized(1), olNormalWindow(2);

  private final int value;

  private OlWindowStateEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlWindowStateEnum findEnum(Integer value) {

    if (value != null) {
      for (OlWindowStateEnum objEnum : values()) {
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

  public boolean isOlMaximized() {

    return olMaximized == this;
  }

  public boolean isOlMinimized() {

    return olMinimized == this;
  }

  public boolean isOlNormalWindow() {

    return olNormalWindow == this;
  }
}
