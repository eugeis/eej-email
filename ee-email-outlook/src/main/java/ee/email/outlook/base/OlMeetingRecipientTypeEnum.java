package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlMeetingRecipientType</a>
 *      </p>
 * @author eugeis
 */
public enum OlMeetingRecipientTypeEnum {
  olOrganizer(0), olRequired(1), olOptional(2), olResource(3);

  private final int value;

  private OlMeetingRecipientTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlMeetingRecipientTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlMeetingRecipientTypeEnum objEnum : values()) {
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

  public boolean isOlOrganizer() {

    return olOrganizer == this;
  }

  public boolean isOlRequired() {

    return olRequired == this;
  }

  public boolean isOlOptional() {

    return olOptional == this;
  }

  public boolean isOlResource() {

    return olResource == this;
  }
}
