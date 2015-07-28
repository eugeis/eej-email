package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlFlagIcon</a>
 *      </p>
 * @author eugeis
 */
public enum OlFlagIconEnum {
  olNoFlagIcon(0), olPurpleFlagIcon(1), olOrangeFlagIcon(2), olGreenFlagIcon(3), olYellowFlagIcon(4), olBlueFlagIcon(5), olRedFlagIcon(
      6);

  private final int value;

  private OlFlagIconEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlFlagIconEnum findEnum(Integer value) {

    if (value != null) {
      for (OlFlagIconEnum objEnum : values()) {
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

  public boolean isOlNoFlagIcon() {

    return olNoFlagIcon == this;
  }

  public boolean isOlPurpleFlagIcon() {

    return olPurpleFlagIcon == this;
  }

  public boolean isOlOrangeFlagIcon() {

    return olOrangeFlagIcon == this;
  }

  public boolean isOlGreenFlagIcon() {

    return olGreenFlagIcon == this;
  }

  public boolean isOlYellowFlagIcon() {

    return olYellowFlagIcon == this;
  }

  public boolean isOlBlueFlagIcon() {

    return olBlueFlagIcon == this;
  }

  public boolean isOlRedFlagIcon() {

    return olRedFlagIcon == this;
  }
}
