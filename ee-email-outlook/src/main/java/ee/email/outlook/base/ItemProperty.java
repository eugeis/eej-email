package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210928(v=office.11).aspx">ItemProperty</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171452(v=office.11).aspx">IsUserProperty</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221795(v=office.11).aspx">Value</a>
 *      </p>*
 * @author eugeis
 */

public class ItemProperty extends OleAuto {

  protected Boolean isUserProperty;

  protected Variant name;

  protected Variant type;

  protected Variant value;

  public ItemProperty(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getIsUserProperty();
    getName();
    getType();
    getValue();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171452(v=office.11).aspx">IsUserProperty</a>
   */
  public Boolean getIsUserProperty() {

    String propertyName = "IsUserProperty";
    try {
      if (this.isUserProperty == null) {
        this.isUserProperty = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isUserProperty;
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

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221795(v=office.11).aspx">Value</a>
   */
  public Variant getValue() {

    String propertyName = "Value";
    try {
      if (this.value == null) {
        this.value = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.value;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
