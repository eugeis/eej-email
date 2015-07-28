package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlActionReplyStyle</a>
 *      </p>
 * @author eugeis
 */
public enum OlActionReplyStyleEnum {
  olOmitOriginalText(0), olEmbedOriginalItem(1), olReplyTickOriginalText(1000), olIncludeOriginalText(2), olIndentOriginalText(
      3), olLinkOriginalItem(4), olUserPreference(5);

  private final int value;

  private OlActionReplyStyleEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlActionReplyStyleEnum findEnum(Integer value) {

    if (value != null) {
      for (OlActionReplyStyleEnum objEnum : values()) {
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

  public boolean isOlOmitOriginalText() {

    return olOmitOriginalText == this;
  }

  public boolean isOlEmbedOriginalItem() {

    return olEmbedOriginalItem == this;
  }

  public boolean isOlReplyTickOriginalText() {

    return olReplyTickOriginalText == this;
  }

  public boolean isOlIncludeOriginalText() {

    return olIncludeOriginalText == this;
  }

  public boolean isOlIndentOriginalText() {

    return olIndentOriginalText == this;
  }

  public boolean isOlLinkOriginalItem() {

    return olLinkOriginalItem == this;
  }

  public boolean isOlUserPreference() {

    return olUserPreference == this;
  }
}
