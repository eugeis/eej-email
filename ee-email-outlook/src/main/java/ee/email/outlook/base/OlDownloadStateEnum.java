package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlDownloadState</a>
 *      </p>
 * @author eugeis
 */
public enum OlDownloadStateEnum {
  olHeaderOnly(0), olFullItem(1);

  private final int value;

  private OlDownloadStateEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlDownloadStateEnum findEnum(Integer value) {

    if (value != null) {
      for (OlDownloadStateEnum objEnum : values()) {
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

  public boolean isOlHeaderOnly() {

    return olHeaderOnly == this;
  }

  public boolean isOlFullItem() {

    return olFullItem == this;
  }
}
