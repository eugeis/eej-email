package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlViewSaveOption</a>
 *      </p>
 * @author eugeis
 */
public enum OlViewSaveOptionEnum {
  olViewSaveOptionThisFolderEveryone(0), olViewSaveOptionThisFolderOnlyMe(1), olViewSaveOptionAllFoldersOfType(2);

  private final int value;

  private OlViewSaveOptionEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlViewSaveOptionEnum findEnum(Integer value) {

    if (value != null) {
      for (OlViewSaveOptionEnum objEnum : values()) {
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

  public boolean isOlViewSaveOptionThisFolderEveryone() {

    return olViewSaveOptionThisFolderEveryone == this;
  }

  public boolean isOlViewSaveOptionThisFolderOnlyMe() {

    return olViewSaveOptionThisFolderOnlyMe == this;
  }

  public boolean isOlViewSaveOptionAllFoldersOfType() {

    return olViewSaveOptionAllFoldersOfType == this;
  }
}
