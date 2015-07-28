package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlRemoteStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlRemoteStatusEnum {
  olRemoteStatusNone(0), olUnMarked(1), olMarkedForDownload(2), olMarkedForCopy(3), olMarkedForDelete(4);

  private final int value;

  private OlRemoteStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlRemoteStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlRemoteStatusEnum objEnum : values()) {
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

  public boolean isOlRemoteStatusNone() {

    return olRemoteStatusNone == this;
  }

  public boolean isOlUnMarked() {

    return olUnMarked == this;
  }

  public boolean isOlMarkedForDownload() {

    return olMarkedForDownload == this;
  }

  public boolean isOlMarkedForCopy() {

    return olMarkedForCopy == this;
  }

  public boolean isOlMarkedForDelete() {

    return olMarkedForDelete == this;
  }
}
