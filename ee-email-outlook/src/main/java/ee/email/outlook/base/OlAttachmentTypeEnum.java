package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlAttachmentType</a>
 *      </p>
 * @author eugeis
 */
public enum OlAttachmentTypeEnum {
  olByValue(1), olByReference(4), olEmbeddeditem(5), olOLE(6);

  private final int value;

  private OlAttachmentTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlAttachmentTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlAttachmentTypeEnum objEnum : values()) {
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

  public boolean isOlByValue() {

    return olByValue == this;
  }

  public boolean isOlByReference() {

    return olByReference == this;
  }

  public boolean isOlEmbeddeditem() {

    return olEmbeddeditem == this;
  }

  public boolean isOlOLE() {

    return olOLE == this;
  }
}
