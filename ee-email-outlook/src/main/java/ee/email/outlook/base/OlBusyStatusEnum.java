package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlBusyStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlBusyStatusEnum {
  olFree(0), olTentative(1), olBusy(2), olOutOfOffice(3);

  private final int value;

  private OlBusyStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlBusyStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlBusyStatusEnum objEnum : values()) {
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

  public boolean isOlFree() {

    return olFree == this;
  }

  public boolean isOlTentative() {

    return olTentative == this;
  }

  public boolean isOlBusy() {

    return olBusy == this;
  }

  public boolean isOlOutOfOffice() {

    return olOutOfOffice == this;
  }
}
