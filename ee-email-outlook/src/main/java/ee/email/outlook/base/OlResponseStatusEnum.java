package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlResponseStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlResponseStatusEnum {
  olResponseNone(0), olResponseOrganized(1), olResponseTentative(2), olResponseAccepted(3), olResponseDeclined(4), olResponseNotResponded(5);

  private final int value;

  private OlResponseStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlResponseStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlResponseStatusEnum objEnum : values()) {
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

  public boolean isOlResponseNone() {

    return olResponseNone == this;
  }

  public boolean isOlResponseOrganized() {

    return olResponseOrganized == this;
  }

  public boolean isOlResponseTentative() {

    return olResponseTentative == this;
  }

  public boolean isOlResponseAccepted() {

    return olResponseAccepted == this;
  }

  public boolean isOlResponseDeclined() {

    return olResponseDeclined == this;
  }

  public boolean isOlResponseNotResponded() {

    return olResponseNotResponded == this;
  }
}
