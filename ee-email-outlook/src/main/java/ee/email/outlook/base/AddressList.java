package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210895(v=office.11).aspx">AddressList</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211353(v=office.11).aspx">AddressEntries</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171429(v=office.11).aspx">ID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171435(v=office.11).aspx">Index</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171448(v=office.11).aspx">IsReadOnly</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210890(v=office.11).aspx">AddressEntries</a>
 *      </p>*
 * @author eugeis
 */

public class AddressList extends OleAuto {

  protected AddressEntries addressEntries;

  protected String iD;

  protected Variant index;

  protected Boolean isReadOnly;

  protected Variant name;

  public AddressList(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getAddressEntries();
    getID();
    getIndex();
    getIsReadOnly();
    getName();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211353(v=office.11).aspx">AddressEntries</a>
   */
  public AddressEntries getAddressEntries() {

    String propertyName = "AddressEntries";
    try {
      if (this.addressEntries == null) {
        this.addressEntries = new AddressEntries(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.addressEntries;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171429(v=office.11).aspx">ID</a>
   */
  public String getID() {

    String propertyName = "ID";
    try {
      if (this.iD == null) {
        this.iD = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.iD;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171448(v=office.11).aspx">IsReadOnly</a>
   */
  public Boolean getIsReadOnly() {

    String propertyName = "IsReadOnly";
    try {
      if (this.isReadOnly == null) {
        this.isReadOnly = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isReadOnly;
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

  @Override
  public void dispose() {

    super.dispose();
    if (this.addressEntries != null) {
      this.addressEntries.dispose();
    }
  }

}
