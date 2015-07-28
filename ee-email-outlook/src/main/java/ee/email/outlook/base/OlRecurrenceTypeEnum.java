package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlRecurrenceType</a>
 *      </p>
 * @author eugeis
 */
public enum OlRecurrenceTypeEnum {
  olRecursDaily(0), olRecursWeekly(1), olRecursMonthly(2), olRecursMonthNth(3), olRecursYearly(5), olRecursYearNth(6);

  private final int value;

  private OlRecurrenceTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlRecurrenceTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlRecurrenceTypeEnum objEnum : values()) {
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

  public boolean isOlRecursDaily() {

    return olRecursDaily == this;
  }

  public boolean isOlRecursWeekly() {

    return olRecursWeekly == this;
  }

  public boolean isOlRecursMonthly() {

    return olRecursMonthly == this;
  }

  public boolean isOlRecursMonthNth() {

    return olRecursMonthNth == this;
  }

  public boolean isOlRecursYearly() {

    return olRecursYearly == this;
  }

  public boolean isOlRecursYearNth() {

    return olRecursYearNth == this;
  }
}
