package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlJournalRecipientType</a>
 *      </p>
 * @author eugeis
 */
public enum OlJournalRecipientTypeEnum {
  olAssociatedContact(1);

  private final int value;

  private OlJournalRecipientTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlJournalRecipientTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlJournalRecipientTypeEnum objEnum : values()) {
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

  public boolean isOlAssociatedContact() {

    return olAssociatedContact == this;
  }
}
