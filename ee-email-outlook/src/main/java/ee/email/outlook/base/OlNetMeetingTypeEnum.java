package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlNetMeetingType</a>
 *      </p>
 * @author eugeis
 */
public enum OlNetMeetingTypeEnum {
  olNetMeeting(0), olNetShow(1), olExchangeConferencing(2);

  private final int value;

  private OlNetMeetingTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlNetMeetingTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlNetMeetingTypeEnum objEnum : values()) {
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

  public boolean isOlNetMeeting() {

    return olNetMeeting == this;
  }

  public boolean isOlNetShow() {

    return olNetShow == this;
  }

  public boolean isOlExchangeConferencing() {

    return olExchangeConferencing == this;
  }
}
