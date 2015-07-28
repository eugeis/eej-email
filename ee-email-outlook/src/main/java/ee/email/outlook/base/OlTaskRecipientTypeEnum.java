package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTaskRecipientType</a>
 *      </p>
 * @author eugeis
 */
public enum OlTaskRecipientTypeEnum {
  olUpdate(2), olFinalStatus(3);

  private final int value;

  private OlTaskRecipientTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTaskRecipientTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTaskRecipientTypeEnum objEnum : values()) {
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

  public boolean isOlUpdate() {

    return olUpdate == this;
  }

  public boolean isOlFinalStatus() {

    return olFinalStatus == this;
  }
}
