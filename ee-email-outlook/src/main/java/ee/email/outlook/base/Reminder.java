package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211029(v=office.11).aspx">Reminder</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211785(v=office.11).aspx">Caption</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171453(v=office.11).aspx">IsVisible</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171456(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171766(v=office.11).aspx">NextReminderDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171787(v=office.11).aspx">OriginalReminderDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220089(v=office.11).aspx">Dismiss</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210302(v=office.11).aspx">Snooze</a>
 *      </p>*
 * @author eugeis
 */

public class Reminder extends OleAuto {

  protected String caption;

  protected Boolean isVisible;

  protected Variant item;

  protected Date nextReminderDate;

  protected Date originalReminderDate;

  public Reminder(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getCaption();
    getIsVisible();
    getItem();
    getNextReminderDate();
    getOriginalReminderDate();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211785(v=office.11).aspx">Caption</a>
   */
  public String getCaption() {

    String propertyName = "Caption";
    try {
      if (this.caption == null) {
        this.caption = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.caption;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171453(v=office.11).aspx">IsVisible</a>
   */
  public Boolean getIsVisible() {

    String propertyName = "IsVisible";
    try {
      if (this.isVisible == null) {
        this.isVisible = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isVisible;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171456(v=office.11).aspx">Item</a>
   */
  public Variant getItem() {

    String propertyName = "Item";
    try {
      if (this.item == null) {
        this.item = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.item;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171766(v=office.11).aspx">NextReminderDate</a>
   */
  public Date getNextReminderDate() {

    String propertyName = "NextReminderDate";
    try {
      if (this.nextReminderDate == null) {
        this.nextReminderDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.nextReminderDate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171787(v=office.11).aspx">OriginalReminderDate</a>
   */
  public Date getOriginalReminderDate() {

    String propertyName = "OriginalReminderDate";
    try {
      if (this.originalReminderDate == null) {
        this.originalReminderDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.originalReminderDate;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
