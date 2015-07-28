package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210891(v=office.11).aspx">AddressEntry</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211363(v=office.11).aspx">Address</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211839(v=office.11).aspx">DisplayType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171429(v=office.11).aspx">ID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171483(v=office.11).aspx">Manager</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171489(v=office.11).aspx">Members</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220087(v=office.11).aspx">Details</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220104(v=office.11).aspx">GetFreeBusy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210315(v=office.11).aspx">Update</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210891(v=office.11).aspx">AddressEntry</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa211006(v=office.11).aspx">Recipient</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210890(v=office.11).aspx">AddressEntries</a>
 *      | <a href="http://msdn.microsoft.com/en-us/library/aa210891(v=office.11).aspx">AddressEntry</a>
 *      </p>*
 * @author eugeis
 */

public class AddressEntry extends OleAuto {

  protected String address;

  protected OlDisplayTypeEnum displayType;

  protected String iD;

  protected Variant manager;

  protected Variant members;

  protected Variant name;

  protected Variant type;

  public AddressEntry(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getAddress();
    getDisplayType();
    getID();
    getManager();
    getMembers();
    getName();
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171483(v=office.11).aspx">Manager</a>
   */
  public Variant getManager() {

    String propertyName = "Manager";
    try {
      if (this.manager == null) {
        this.manager = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.manager;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171489(v=office.11).aspx">Members</a>
   */
  public Variant getMembers() {

    String propertyName = "Members";
    try {
      if (this.members == null) {
        this.members = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.members;
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
  }

}
