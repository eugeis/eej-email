package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlConnectionMode</a>
 *      </p>
 * @author eugeis
 */
public enum OlConnectionModeEnum {
  olOffline(100), olLowBandwidth(200), olOnline(300);

  private final int value;

  private OlConnectionModeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlConnectionModeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlConnectionModeEnum objEnum : values()) {
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

  public boolean isOlOffline() {

    return olOffline == this;
  }

  public boolean isOlLowBandwidth() {

    return olLowBandwidth == this;
  }

  public boolean isOlOnline() {

    return olOnline == this;
  }
}
