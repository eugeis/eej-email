package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlExchangeConnectionMode</a>
 *      </p>
 * @author eugeis
 */
public enum OlExchangeConnectionModeEnum {
  olNoExchange(0), olCachedConnectedFull(700), olOffline(100), olOnline(800), olCachedOffline(200), olCachedConnectedHeaders(500), olCachedConnectedDrizzle(600), olCachedDisconnected(400), olDisconnected(300);

  private final int value;

  private OlExchangeConnectionModeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlExchangeConnectionModeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlExchangeConnectionModeEnum objEnum : values()) {
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

  public boolean isOlNoExchange() {

    return olNoExchange == this;
  }

  public boolean isOlCachedConnectedFull() {

    return olCachedConnectedFull == this;
  }

  public boolean isOlOffline() {

    return olOffline == this;
  }

  public boolean isOlOnline() {

    return olOnline == this;
  }

  public boolean isOlCachedOffline() {

    return olCachedOffline == this;
  }

  public boolean isOlCachedConnectedHeaders() {

    return olCachedConnectedHeaders == this;
  }

  public boolean isOlCachedConnectedDrizzle() {

    return olCachedConnectedDrizzle == this;
  }

  public boolean isOlCachedDisconnected() {

    return olCachedDisconnected == this;
  }

  public boolean isOlDisconnected() {

    return olDisconnected == this;
  }
}
