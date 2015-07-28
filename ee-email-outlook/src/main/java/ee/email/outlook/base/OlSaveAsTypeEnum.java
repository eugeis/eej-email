package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlSaveAsType</a>
 *      </p>
 * @author eugeis
 */
public enum OlSaveAsTypeEnum {
  olTXT(0), olRTF(1), olTemplate(2), olMSG(3), olDoc(4), olHTML(5), olVCard(6), olVCal(7), olICal(8), olMSGUnicode(9);

  private final int value;

  private OlSaveAsTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlSaveAsTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlSaveAsTypeEnum objEnum : values()) {
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

  public boolean isOlTXT() {

    return olTXT == this;
  }

  public boolean isOlRTF() {

    return olRTF == this;
  }

  public boolean isOlTemplate() {

    return olTemplate == this;
  }

  public boolean isOlMSG() {

    return olMSG == this;
  }

  public boolean isOlDoc() {

    return olDoc == this;
  }

  public boolean isOlHTML() {

    return olHTML == this;
  }

  public boolean isOlVCard() {

    return olVCard == this;
  }

  public boolean isOlVCal() {

    return olVCal == this;
  }

  public boolean isOlICal() {

    return olICal == this;
  }

  public boolean isOlMSGUnicode() {

    return olMSGUnicode == this;
  }
}
