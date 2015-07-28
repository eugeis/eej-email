package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlMailRecipientType</a>
 *      </p>
 * @author eugeis
 */
public enum OlMailRecipientTypeEnum {
  olOriginator(0), olTo(1), olCC(2), olBCC(3);

  private final int value;

  private OlMailRecipientTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlMailRecipientTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlMailRecipientTypeEnum objEnum : values()) {
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

  public boolean isOlOriginator() {

    return olOriginator == this;
  }

  public boolean isOlTo() {

    return olTo == this;
  }

  public boolean isOlCC() {

    return olCC == this;
  }

  public boolean isOlBCC() {

    return olBCC == this;
  }
}
