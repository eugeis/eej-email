package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlShowItemCount</a>
 *      </p>
 * @author eugeis
 */
public enum OlShowItemCountEnum {
  olNoItemCount(0), olShowUnreadItemCount(1), olShowTotalItemCount(2);

  private final int value;

  private OlShowItemCountEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlShowItemCountEnum findEnum(Integer value) {

    if (value != null) {
      for (OlShowItemCountEnum objEnum : values()) {
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

  public boolean isOlNoItemCount() {

    return olNoItemCount == this;
  }

  public boolean isOlShowUnreadItemCount() {

    return olShowUnreadItemCount == this;
  }

  public boolean isOlShowTotalItemCount() {

    return olShowTotalItemCount == this;
  }
}
