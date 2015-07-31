package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210962(v=office.11).aspx">OutlookBarGroup</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172024(v=office.11).aspx">Shortcuts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221814(v=office.11).aspx">ViewType</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210966(v=office.11).aspx">OutlookBarPane</a>
 *      </p>*
 * @author eugeis
 */

public class OutlookBarGroup extends OleAuto {

  protected Variant name;

  protected Variant shortcuts;

  protected Variant viewType;

  public OutlookBarGroup(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getName();
    getShortcuts();
    getViewType();
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172024(v=office.11).aspx">Shortcuts</a>
   */
  public Variant getShortcuts() {

    String propertyName = "Shortcuts";
    try {
      if (this.shortcuts == null) {
        this.shortcuts = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.shortcuts;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221814(v=office.11).aspx">ViewType</a>
   */
  public Variant getViewType() {

    String propertyName = "ViewType";
    try {
      if (this.viewType == null) {
        this.viewType = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.viewType;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
