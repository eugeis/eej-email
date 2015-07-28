package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlOutlookBarViewType</a>
 *      </p>
 * @author eugeis
 */
public enum OlOutlookBarViewTypeEnum {
  olLargeIcon(0), olSmallIcon(1);

  private final int value;

  private OlOutlookBarViewTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlOutlookBarViewTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlOutlookBarViewTypeEnum objEnum : values()) {
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

  public boolean isOlLargeIcon() {

    return olLargeIcon == this;
  }

  public boolean isOlSmallIcon() {

    return olSmallIcon == this;
  }
}
