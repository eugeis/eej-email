package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlDisplayType</a>
 *      </p>
 * @author eugeis
 */
public enum OlDisplayTypeEnum {
  olUser(0), olDistList(1), olForum(2), olAgent(3), olOrganization(4), olPrivateDistList(5), olRemoteUser(6);

  private final int value;

  private OlDisplayTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlDisplayTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlDisplayTypeEnum objEnum : values()) {
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

  public boolean isOlUser() {

    return olUser == this;
  }

  public boolean isOlDistList() {

    return olDistList == this;
  }

  public boolean isOlForum() {

    return olForum == this;
  }

  public boolean isOlAgent() {

    return olAgent == this;
  }

  public boolean isOlOrganization() {

    return olOrganization == this;
  }

  public boolean isOlPrivateDistList() {

    return olPrivateDistList == this;
  }

  public boolean isOlRemoteUser() {

    return olRemoteUser == this;
  }
}
