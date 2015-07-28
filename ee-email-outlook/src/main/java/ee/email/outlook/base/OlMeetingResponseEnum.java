package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlMeetingResponse</a>
 *      </p>
 * @author eugeis
 */
public enum OlMeetingResponseEnum {
  olMeetingTentative(2), olMeetingAccepted(3), olMeetingDeclined(4);

  private final int value;

  private OlMeetingResponseEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlMeetingResponseEnum findEnum(Integer value) {

    if (value != null) {
      for (OlMeetingResponseEnum objEnum : values()) {
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

  public boolean isOlMeetingTentative() {

    return olMeetingTentative == this;
  }

  public boolean isOlMeetingAccepted() {

    return olMeetingAccepted == this;
  }

  public boolean isOlMeetingDeclined() {

    return olMeetingDeclined == this;
  }
}
