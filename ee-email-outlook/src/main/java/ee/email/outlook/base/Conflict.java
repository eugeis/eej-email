package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210906(v=office.11).aspx">Conflict</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171456(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
 *      </p>*
 * @author eugeis
 */

public class Conflict extends OleAuto {

  protected Variant item;

  protected Variant name;

  protected Variant type;

  public Conflict(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getItem();
    getName();
    getType();
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
