package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlFolderDisplayMode</a>
 *      </p>
 * @author eugeis
 */
public enum OlFolderDisplayModeEnum {
  olFolderDisplayNormal(0), olFolderDisplayFolderOnly(1), olFolderDisplayNoNavigation(2);

  private final int value;

  private OlFolderDisplayModeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlFolderDisplayModeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlFolderDisplayModeEnum objEnum : values()) {
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

  public boolean isOlFolderDisplayNormal() {

    return olFolderDisplayNormal == this;
  }

  public boolean isOlFolderDisplayFolderOnly() {

    return olFolderDisplayFolderOnly == this;
  }

  public boolean isOlFolderDisplayNoNavigation() {

    return olFolderDisplayNoNavigation == this;
  }
}
