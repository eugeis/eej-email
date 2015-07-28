package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlMeetingStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlMeetingStatusEnum {
  olNonMeeting(0), olMeeting(1), olMeetingReceived(3), olMeetingCanceled(5);

  private final int value;

  private OlMeetingStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlMeetingStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlMeetingStatusEnum objEnum : values()) {
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

  public boolean isOlNonMeeting() {

    return olNonMeeting == this;
  }

  public boolean isOlMeeting() {

    return olMeeting == this;
  }

  public boolean isOlMeetingReceived() {

    return olMeetingReceived == this;
  }

  public boolean isOlMeetingCanceled() {

    return olMeetingCanceled == this;
  }
}
