package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTaskStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlTaskStatusEnum {
  olTaskNotStarted(0), olTaskInProgress(1), olTaskComplete(2), olTaskWaiting(3), olTaskDeferred(4);

  private final int value;

  private OlTaskStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTaskStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTaskStatusEnum objEnum : values()) {
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

  public boolean isOlTaskNotStarted() {

    return olTaskNotStarted == this;
  }

  public boolean isOlTaskInProgress() {

    return olTaskInProgress == this;
  }

  public boolean isOlTaskComplete() {

    return olTaskComplete == this;
  }

  public boolean isOlTaskWaiting() {

    return olTaskWaiting == this;
  }

  public boolean isOlTaskDeferred() {

    return olTaskDeferred == this;
  }
}
