package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlInspectorClose</a>
 *      </p>
 * @author eugeis
 */
public enum OlInspectorCloseEnum {
  olSave(0), olDiscard(1), olPromptForSave(2);

  private final int value;

  private OlInspectorCloseEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlInspectorCloseEnum findEnum(Integer value) {

    if (value != null) {
      for (OlInspectorCloseEnum objEnum : values()) {
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

  public boolean isOlSave() {

    return olSave == this;
  }

  public boolean isOlDiscard() {

    return olDiscard == this;
  }

  public boolean isOlPromptForSave() {

    return olPromptForSave == this;
  }
}
