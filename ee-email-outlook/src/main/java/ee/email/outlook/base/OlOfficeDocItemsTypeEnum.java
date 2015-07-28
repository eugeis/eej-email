package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlOfficeDocItemsType</a>
 *      </p>
 * @author eugeis
 */
public enum OlOfficeDocItemsTypeEnum {
  olExcelWorkSheetItem(8), olWordDocumentItem(9), olPowerPointShowItem(10);

  private final int value;

  private OlOfficeDocItemsTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlOfficeDocItemsTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlOfficeDocItemsTypeEnum objEnum : values()) {
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

  public boolean isOlExcelWorkSheetItem() {

    return olExcelWorkSheetItem == this;
  }

  public boolean isOlWordDocumentItem() {

    return olWordDocumentItem == this;
  }

  public boolean isOlPowerPointShowItem() {

    return olPowerPointShowItem == this;
  }
}
