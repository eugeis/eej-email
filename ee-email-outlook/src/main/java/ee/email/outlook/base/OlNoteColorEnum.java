package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlNoteColor</a>
 *      </p>
 * @author eugeis
 */
public enum OlNoteColorEnum {
  olBlue(0), olGreen(1), olPink(2), olYellow(3), olWhite(4);

  private final int value;

  private OlNoteColorEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlNoteColorEnum findEnum(Integer value) {

    if (value != null) {
      for (OlNoteColorEnum objEnum : values()) {
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

  public boolean isOlBlue() {

    return olBlue == this;
  }

  public boolean isOlGreen() {

    return olGreen == this;
  }

  public boolean isOlPink() {

    return olPink == this;
  }

  public boolean isOlYellow() {

    return olYellow == this;
  }

  public boolean isOlWhite() {

    return olWhite == this;
  }
}
