package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlActionCopyLike</a>
 *      </p>
 * @author eugeis
 */
public enum OlActionCopyLikeEnum {
  olReply(0), olReplyAll(1), olForward(2), olReplyFolder(3), olRespond(4);

  private final int value;

  private OlActionCopyLikeEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlActionCopyLikeEnum findEnum(Integer value) {

    if (value != null) {
      for (OlActionCopyLikeEnum objEnum : values()) {
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

  public boolean isOlReply() {

    return olReply == this;
  }

  public boolean isOlReplyAll() {

    return olReplyAll == this;
  }

  public boolean isOlForward() {

    return olForward == this;
  }

  public boolean isOlReplyFolder() {

    return olReplyFolder == this;
  }

  public boolean isOlRespond() {

    return olRespond == this;
  }
}
