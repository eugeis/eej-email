package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTaskOwnership</a>
 *      </p>
 * @author eugeis
 */
public enum OlTaskOwnershipEnum {
  olNewTask(0), olDelegatedTask(1), olOwnTask(2);

  private final int value;

  private OlTaskOwnershipEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTaskOwnershipEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTaskOwnershipEnum objEnum : values()) {
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

  public boolean isOlNewTask() {

    return olNewTask == this;
  }

  public boolean isOlDelegatedTask() {

    return olDelegatedTask == this;
  }

  public boolean isOlOwnTask() {

    return olOwnTask == this;
  }
}
