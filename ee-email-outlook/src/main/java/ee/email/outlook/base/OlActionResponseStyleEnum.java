package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlActionResponseStyle</a>
 *      </p>
 * @author eugeis
 */
public enum OlActionResponseStyleEnum {
  olOpen(0), olSend(1), olPrompt(2);

  private final int value;

  private OlActionResponseStyleEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlActionResponseStyleEnum findEnum(Integer value) {

    if (value != null) {
      for (OlActionResponseStyleEnum objEnum : values()) {
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

  public boolean isOlOpen() {

    return olOpen == this;
  }

  public boolean isOlSend() {

    return olSend == this;
  }

  public boolean isOlPrompt() {

    return olPrompt == this;
  }
}
