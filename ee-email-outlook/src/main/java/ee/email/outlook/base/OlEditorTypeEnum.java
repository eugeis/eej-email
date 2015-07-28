package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlEditorType</a>
 *      </p>
 * @author eugeis
 */
public enum OlEditorTypeEnum {
  olEditorText(1), olEditorHTML(2), olEditorRTF(3), olEditorWord(4);

  private final int value;

  private OlEditorTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlEditorTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlEditorTypeEnum objEnum : values()) {
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

  public boolean isOlEditorText() {

    return olEditorText == this;
  }

  public boolean isOlEditorHTML() {

    return olEditorHTML == this;
  }

  public boolean isOlEditorRTF() {

    return olEditorRTF == this;
  }

  public boolean isOlEditorWord() {

    return olEditorWord == this;
  }
}
