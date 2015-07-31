package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211006(v=office.11).aspx">Recipient</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211363(v=office.11).aspx">Address</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211357(v=office.11).aspx">AddressEntry</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211417(v=office.11).aspx">AutoResponse</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211839(v=office.11).aspx">DisplayType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171435(v=office.11).aspx">Index</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171485(v=office.11).aspx">MeetingResponseStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171917(v=office.11).aspx">Resolved</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220710(v=office.11).aspx">TrackingStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220697(v=office.11).aspx">TrackingStatusTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220097(v=office.11).aspx">FreeBusy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210271(v=office.11).aspx">Resolve</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210954(v=office.11).aspx">NameSpace</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210891(v=office.11).aspx">AddressEntry</a>
 *      </p>*
 * @author eugeis
 */

public class Recipient extends OleAuto {

  protected String address;

  protected AddressEntry addressEntry;

  protected String autoResponse;

  protected OlDisplayTypeEnum displayType;

  protected String entryID;

  protected Variant index;

  protected OlResponseStatusEnum meetingResponseStatus;

  protected Variant name;

  protected Boolean resolved;

  protected OlTrackingStatusEnum trackingStatus;

  protected Date trackingStatusTime;

  protected Variant type;

  public Recipient(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getAddress();
    getAddressEntry();
    getAutoResponse();
    getDisplayType();
    getEntryID();
    getIndex();
    getMeetingResponseStatus();
    getName();
    getResolved();
    getTrackingStatus();
    getTrackingStatusTime();
    getType();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211363(v=office.11).aspx">Address</a>
   */
  public String getAddress() {

    String propertyName = "Address";
    try {
      if (this.address == null) {
        this.address = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.address;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211357(v=office.11).aspx">AddressEntry</a>
   */
  public AddressEntry getAddressEntry() {

    String propertyName = "AddressEntry";
    try {
      if (this.addressEntry == null) {
        this.addressEntry = new AddressEntry(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.addressEntry;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211417(v=office.11).aspx">AutoResponse</a>
   */
  public String getAutoResponse() {

    String propertyName = "AutoResponse";
    try {
      if (this.autoResponse == null) {
        this.autoResponse = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.autoResponse;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211839(v=office.11).aspx">DisplayType</a>
   */
  public OlDisplayTypeEnum getDisplayType() {

    String propertyName = "DisplayType";
    try {
      if (this.displayType == null) {
        this.displayType = OlDisplayTypeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.displayType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a>
   */
  public String getEntryID() {

    String propertyName = "EntryID";
    try {
      if (this.entryID == null) {
        this.entryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.entryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171435(v=office.11).aspx">Index</a>
   */
  public Variant getIndex() {

    String propertyName = "Index";
    try {
      if (this.index == null) {
        this.index = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.index;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171485(v=office.11).aspx">MeetingResponseStatus</a>
   */
  public OlResponseStatusEnum getMeetingResponseStatus() {

    String propertyName = "MeetingResponseStatus";
    try {
      if (this.meetingResponseStatus == null) {
        this.meetingResponseStatus = OlResponseStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.meetingResponseStatus;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a>
   */
  public Variant getName() {

    String propertyName = "Name";
    try {
      if (this.name == null) {
        this.name = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.name;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171917(v=office.11).aspx">Resolved</a>
   */
  public Boolean getResolved() {

    String propertyName = "Resolved";
    try {
      if (this.resolved == null) {
        this.resolved = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.resolved;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220710(v=office.11).aspx">TrackingStatus</a>
   */
  public OlTrackingStatusEnum getTrackingStatus() {

    String propertyName = "TrackingStatus";
    try {
      if (this.trackingStatus == null) {
        this.trackingStatus = OlTrackingStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.trackingStatus;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220697(v=office.11).aspx">TrackingStatusTime</a>
   */
  public Date getTrackingStatusTime() {

    String propertyName = "TrackingStatusTime";
    try {
      if (this.trackingStatusTime == null) {
        this.trackingStatusTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.trackingStatusTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
   */
  public Variant getType() {

    String propertyName = "Type";
    try {
      if (this.type == null) {
        this.type = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.type;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.addressEntry != null) {
      this.addressEntry.dispose();
    }
  }

}
