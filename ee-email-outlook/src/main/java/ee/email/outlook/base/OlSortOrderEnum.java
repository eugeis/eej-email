package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlSortOrder</a>
 *      </p>
 * @author eugeis
 */
public enum OlSortOrderEnum {
  olSortNone(0), olAscending(1), olDescending(2);

  private final int value;

  private OlSortOrderEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlSortOrderEnum findEnum(Integer value) {

    if (value != null) {
      for (OlSortOrderEnum objEnum : values()) {
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

  public boolean isOlSortNone() {

    return olSortNone == this;
  }

  public boolean isOlAscending() {

    return olAscending == this;
  }

  public boolean isOlDescending() {

    return olDescending == this;
  }
}
