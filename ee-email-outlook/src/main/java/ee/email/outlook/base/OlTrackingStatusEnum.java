package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlTrackingStatus</a>
 *      </p>
 * @author eugeis
 */
public enum OlTrackingStatusEnum {
  olTrackingNone(0), olTrackingDelivered(1), olTrackingNotDelivered(2), olTrackingNotRead(3), olTrackingRecallFailure(4), olTrackingRecallSuccess(5), olTrackingRead(6), olTrackingReplied(7);

  private final int value;

  private OlTrackingStatusEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlTrackingStatusEnum findEnum(Integer value) {

    if (value != null) {
      for (OlTrackingStatusEnum objEnum : values()) {
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

  public boolean isOlTrackingNone() {

    return olTrackingNone == this;
  }

  public boolean isOlTrackingDelivered() {

    return olTrackingDelivered == this;
  }

  public boolean isOlTrackingNotDelivered() {

    return olTrackingNotDelivered == this;
  }

  public boolean isOlTrackingNotRead() {

    return olTrackingNotRead == this;
  }

  public boolean isOlTrackingRecallFailure() {

    return olTrackingRecallFailure == this;
  }

  public boolean isOlTrackingRecallSuccess() {

    return olTrackingRecallSuccess == this;
  }

  public boolean isOlTrackingRead() {

    return olTrackingRead == this;
  }

  public boolean isOlTrackingReplied() {

    return olTrackingReplied == this;
  }
}
