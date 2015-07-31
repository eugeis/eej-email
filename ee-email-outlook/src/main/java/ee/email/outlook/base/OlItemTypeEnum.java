package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlItemType</a>
 *      </p>
 * @author eugeis
 */
public enum OlItemTypeEnum {
  olMailItem(0), olAppointmentItem(1), olContactItem(2), olTaskItem(3), olJournalItem(4), olNoteItem(5), olPostItem(6), olDistributionListItem(7);

  private final int value;

  private OlItemTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlItemTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlItemTypeEnum objEnum : values()) {
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

  public boolean isOlMailItem() {

    return olMailItem == this;
  }

  public boolean isOlAppointmentItem() {

    return olAppointmentItem == this;
  }

  public boolean isOlContactItem() {

    return olContactItem == this;
  }

  public boolean isOlTaskItem() {

    return olTaskItem == this;
  }

  public boolean isOlJournalItem() {

    return olJournalItem == this;
  }

  public boolean isOlNoteItem() {

    return olNoteItem == this;
  }

  public boolean isOlPostItem() {

    return olPostItem == this;
  }

  public boolean isOlDistributionListItem() {

    return olDistributionListItem == this;
  }
}
