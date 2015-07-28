package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlPane</a>
 *      </p>
 * @author eugeis
 */
public enum OlPaneEnum {
  olOutlookBar(1), olFolderList(2), olPreview(3), olNavigationPane(4);

  private final int value;

  private OlPaneEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlPaneEnum findEnum(Integer value) {

    if (value != null) {
      for (OlPaneEnum objEnum : values()) {
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

  public boolean isOlOutlookBar() {

    return olOutlookBar == this;
  }

  public boolean isOlFolderList() {

    return olFolderList == this;
  }

  public boolean isOlPreview() {

    return olPreview == this;
  }

  public boolean isOlNavigationPane() {

    return olNavigationPane == this;
  }
}
