package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210913(v=office.11).aspx">Exception</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211396(v=office.11).aspx">AppointmentItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211834(v=office.11).aspx">Deleted</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171786(v=office.11).aspx">OriginalDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210899(v=office.11).aspx">AppointmentItem</a>
 *      | <a href="http://msdn.microsoft.com/en-us/library/aa210924(v=office.11).aspx">ItemProperties</a>
 *      </p>*
 * @author eugeis
 */

public class OleException extends OleAuto {

  protected AppointmentItem appointmentItem;

  protected Boolean deleted;

  protected ItemProperties itemProperties;

  protected Date originalDate;

  public OleException(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getAppointmentItem();
    getDeleted();
    getItemProperties();
    getOriginalDate();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211396(v=office.11).aspx">AppointmentItem</a>
   */
  public AppointmentItem getAppointmentItem() {

    String propertyName = "AppointmentItem";
    try {
      if (this.appointmentItem == null) {
        this.appointmentItem = new AppointmentItem(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.appointmentItem;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211834(v=office.11).aspx">Deleted</a>
   */
  public Boolean getDeleted() {

    String propertyName = "Deleted";
    try {
      if (this.deleted == null) {
        this.deleted = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.deleted;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a>
   */
  public ItemProperties getItemProperties() {

    String propertyName = "ItemProperties";
    try {
      if (this.itemProperties == null) {
        this.itemProperties = new ItemProperties(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.itemProperties;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171786(v=office.11).aspx">OriginalDate</a>
   */
  public Date getOriginalDate() {

    String propertyName = "OriginalDate";
    try {
      if (this.originalDate == null) {
        this.originalDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.originalDate;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.appointmentItem != null) {
      this.appointmentItem.dispose();
    }
    if (this.itemProperties != null) {
      this.itemProperties.dispose();
    }
  }

}
