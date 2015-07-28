package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlActionShowOn</a>
 *      </p>
 * @author eugeis
 */
public enum OlActionShowOnEnum {
  olDontShow(0), olMenu(1), olMenuAndToolbar(2);

  private final int value;

  private OlActionShowOnEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlActionShowOnEnum findEnum(Integer value) {

    if (value != null) {
      for (OlActionShowOnEnum objEnum : values()) {
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

  public boolean isOlDontShow() {

    return olDontShow == this;
  }

  public boolean isOlMenu() {

    return olMenu == this;
  }

  public boolean isOlMenuAndToolbar() {

    return olMenuAndToolbar == this;
  }
}
