package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTaskDelegationState</a>
 *      </p>
 * @author eugeis
 */
public enum OlTaskDelegationStateEnum {
  olTaskNotDelegated(0), olTaskDelegationUnknown(1), olTaskDelegationAccepted(2), olTaskDelegationDeclined(3);

  private final int value;

  private OlTaskDelegationStateEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTaskDelegationStateEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTaskDelegationStateEnum objEnum : values()) {
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

  public boolean isOlTaskNotDelegated() {

    return olTaskNotDelegated == this;
  }

  public boolean isOlTaskDelegationUnknown() {

    return olTaskDelegationUnknown == this;
  }

  public boolean isOlTaskDelegationAccepted() {

    return olTaskDelegationAccepted == this;
  }

  public boolean isOlTaskDelegationDeclined() {

    return olTaskDelegationDeclined == this;
  }
}
