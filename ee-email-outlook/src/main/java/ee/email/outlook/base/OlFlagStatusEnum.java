package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlFlagStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlFlagStatusEnum {
  olNoFlag(0), olFlagComplete(1), olFlagMarked(2);

  private final int value;

  private OlFlagStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlFlagStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlFlagStatusEnum objEnum : values()) {
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

  public boolean isOlNoFlag() {

    return olNoFlag == this;
  }

  public boolean isOlFlagComplete() {

    return olFlagComplete == this;
  }

  public boolean isOlFlagMarked() {

    return olFlagMarked == this;
  }
}
