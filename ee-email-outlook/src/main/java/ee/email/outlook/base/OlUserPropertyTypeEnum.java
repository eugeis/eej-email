package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlUserPropertyType</a>
 *      </p>
 * @author eugeis
 */
public enum OlUserPropertyTypeEnum {
  olOutlookInternal(0), olText(1), olCombination(19), olFormula(18), olNumber(3), olDateTime(5), olYesNo(6), olDuration(7), olKeywords(11), olPercent(12), olCurrency(14);

  private final int value;

  private OlUserPropertyTypeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlUserPropertyTypeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlUserPropertyTypeEnum objEnum : values()) {
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

  public boolean isOlOutlookInternal() {

    return olOutlookInternal == this;
  }

  public boolean isOlText() {

    return olText == this;
  }

  public boolean isOlCombination() {

    return olCombination == this;
  }

  public boolean isOlFormula() {

    return olFormula == this;
  }

  public boolean isOlNumber() {

    return olNumber == this;
  }

  public boolean isOlDateTime() {

    return olDateTime == this;
  }

  public boolean isOlYesNo() {

    return olYesNo == this;
  }

  public boolean isOlDuration() {

    return olDuration == this;
  }

  public boolean isOlKeywords() {

    return olKeywords == this;
  }

  public boolean isOlPercent() {

    return olPercent == this;
  }

  public boolean isOlCurrency() {

    return olCurrency == this;
  }
}
