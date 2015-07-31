package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211053(v=office.11).aspx">Selection</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210916(v=office.11).aspx">Explorer</a>
 *      </p>*
 * @author eugeis
 */

public class Selection extends OleAuto {

  protected Variant count;

  public Selection(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getCount();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a>
   */
  public Variant getCount() {

    String propertyName = "Count";
    try {
      if (this.count == null) {
        this.count = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.count;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
