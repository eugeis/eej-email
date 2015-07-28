package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlDaysOfWeek</a>
 *      </p>
 * @author eugeis
 */
public enum OlDaysOfWeekEnum {
  olSunday(1), olThursday(16), olFriday(32), olMonday(2), olSaturday(64), olTuesday(4), olWednesday(8);

  private final int value;

  private OlDaysOfWeekEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlDaysOfWeekEnum findEnum(Integer value) {

    if (value != null) {
      for (OlDaysOfWeekEnum objEnum : values()) {
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

  public boolean isOlSunday() {

    return olSunday == this;
  }

  public boolean isOlThursday() {

    return olThursday == this;
  }

  public boolean isOlFriday() {

    return olFriday == this;
  }

  public boolean isOlMonday() {

    return olMonday == this;
  }

  public boolean isOlSaturday() {

    return olSaturday == this;
  }

  public boolean isOlTuesday() {

    return olTuesday == this;
  }

  public boolean isOlWednesday() {

    return olWednesday == this;
  }
}
