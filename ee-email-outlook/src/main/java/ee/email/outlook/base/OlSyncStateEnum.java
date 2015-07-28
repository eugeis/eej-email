package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlSyncState</a>
 *      </p>
 * @author eugeis
 */
public enum OlSyncStateEnum {
  olSyncStopped(0), olSyncStarted(1);

  private final int value;

  private OlSyncStateEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlSyncStateEnum findEnum(Integer value) {

    if (value != null) {
      for (OlSyncStateEnum objEnum : values()) {
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

  public boolean isOlSyncStopped() {

    return olSyncStopped == this;
  }

  public boolean isOlSyncStarted() {

    return olSyncStarted == this;
  }
}
