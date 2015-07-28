package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlRecurrenceState</a>
 *      </p>
 * @author eugeis
 */
public enum OlRecurrenceStateEnum {
  olApptNotRecurring(0), olApptMaster(1), olApptOccurrence(2), olApptException(3);

  private final int value;

  private OlRecurrenceStateEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlRecurrenceStateEnum findEnum(Integer value) {

    if (value != null) {
      for (OlRecurrenceStateEnum objEnum : values()) {
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

  public boolean isOlApptNotRecurring() {

    return olApptNotRecurring == this;
  }

  public boolean isOlApptMaster() {

    return olApptMaster == this;
  }

  public boolean isOlApptOccurrence() {

    return olApptOccurrence == this;
  }

  public boolean isOlApptException() {

    return olApptException == this;
  }
}
