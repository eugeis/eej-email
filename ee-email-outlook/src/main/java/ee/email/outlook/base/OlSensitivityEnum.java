package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlSensitivity</a>
 *      </p>
 * @author eugeis
 */
public enum OlSensitivityEnum {
  olNormal(0), olPersonal(1), olPrivate(2), olConfidential(3);

  private final int value;

  private OlSensitivityEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlSensitivityEnum findEnum(Integer value) {

    if (value != null) {
      for (OlSensitivityEnum objEnum : values()) {
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

  public boolean isOlNormal() {

    return olNormal == this;
  }

  public boolean isOlPersonal() {

    return olPersonal == this;
  }

  public boolean isOlPrivate() {

    return olPrivate == this;
  }

  public boolean isOlConfidential() {

    return olConfidential == this;
  }
}
