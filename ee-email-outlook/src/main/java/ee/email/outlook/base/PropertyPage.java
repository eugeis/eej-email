package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210988(v=office.11).aspx">PropertyPage</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211837(v=office.11).aspx">Dirty</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220072(v=office.11).aspx">Apply</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220112(v=office.11).aspx">GetPageInfo</a>
 *      </p>*
 * @author eugeis
 */

public class PropertyPage extends OleAuto {

  protected Boolean dirty;

  public PropertyPage(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getDirty();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211837(v=office.11).aspx">Dirty</a>
   */
  public Boolean getDirty() {

    String propertyName = "Dirty";
    try {
      if (this.dirty == null) {
        this.dirty = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.dirty;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
