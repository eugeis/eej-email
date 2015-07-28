package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlGender</a>
 *      </p>
 * @author eugeis
 */
public enum OlGenderEnum {
  olUnspecified(0), olFemale(1), olMale(2);

  private final int value;

  private OlGenderEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlGenderEnum findEnum(Integer value) {

    if (value != null) {
      for (OlGenderEnum objEnum : values()) {
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

  public boolean isOlUnspecified() {

    return olUnspecified == this;
  }

  public boolean isOlFemale() {

    return olFemale == this;
  }

  public boolean isOlMale() {

    return olMale == this;
  }
}
