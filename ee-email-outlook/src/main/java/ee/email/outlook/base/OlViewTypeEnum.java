package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlViewType</a>
 *      </p>
 * @author eugeis
 */
public enum OlViewTypeEnum {
  olTableView(0), olCardView(1), olCalendarView(2), olIconView(3), olTimelineView(4);

  private final int value;

  private OlViewTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlViewTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlViewTypeEnum objEnum : values()) {
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

  public boolean isOlTableView() {

    return olTableView == this;
  }

  public boolean isOlCardView() {

    return olCardView == this;
  }

  public boolean isOlCalendarView() {

    return olCalendarView == this;
  }

  public boolean isOlIconView() {

    return olIconView == this;
  }

  public boolean isOlTimelineView() {

    return olTimelineView == this;
  }
}
