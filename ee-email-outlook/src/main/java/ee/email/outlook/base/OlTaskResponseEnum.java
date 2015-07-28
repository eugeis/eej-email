package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTaskResponse</a>
 *      </p>
 * @author eugeis
 */
public enum OlTaskResponseEnum {
  olTaskSimple(0), olTaskAssign(1), olTaskAccept(2), olTaskDecline(3);

  private final int value;

  private OlTaskResponseEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTaskResponseEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTaskResponseEnum objEnum : values()) {
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

  public boolean isOlTaskSimple() {

    return olTaskSimple == this;
  }

  public boolean isOlTaskAssign() {

    return olTaskAssign == this;
  }

  public boolean isOlTaskAccept() {

    return olTaskAccept == this;
  }

  public boolean isOlTaskDecline() {

    return olTaskDecline == this;
  }
}
