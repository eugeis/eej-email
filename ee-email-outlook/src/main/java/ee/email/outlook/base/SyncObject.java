package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211062(v=office.11).aspx">SyncObject</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa210308(v=office.11).aspx">Start</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210314(v=office.11).aspx">Stop</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171312(v=office.11).aspx">OnError</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171324(v=office.11).aspx">Progress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219364(v=office.11).aspx">SyncEnd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219365(v=office.11).aspx">SyncStart</a>
 *      </p>*
 * @author eugeis
 */

public class SyncObject extends OleAuto {

  protected Variant name;

  public SyncObject(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getName();
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
  }

}
