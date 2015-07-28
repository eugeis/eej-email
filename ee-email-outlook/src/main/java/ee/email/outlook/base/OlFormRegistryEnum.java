package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlFormRegistry</a>
 *      </p>
 * @author eugeis
 */
public enum OlFormRegistryEnum {
  olDefaultRegistry(0), olPersonalRegistry(2), olFolderRegistry(3), olOrganizationRegistry(4);

  private final int value;

  private OlFormRegistryEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlFormRegistryEnum findEnum(Integer value) {

    if (value != null) {
      for (OlFormRegistryEnum objEnum : values()) {
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

  public boolean isOlDefaultRegistry() {

    return olDefaultRegistry == this;
  }

  public boolean isOlPersonalRegistry() {

    return olPersonalRegistry == this;
  }

  public boolean isOlFolderRegistry() {

    return olFolderRegistry == this;
  }

  public boolean isOlOrganizationRegistry() {

    return olOrganizationRegistry == this;
  }
}
