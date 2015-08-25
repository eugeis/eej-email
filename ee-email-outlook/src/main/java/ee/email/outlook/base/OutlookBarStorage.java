package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210976(v=office.11).aspx">OutlookBarStorage</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212049(v=office.11).aspx">Groups</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210966(v=office.11).aspx">OutlookBarPane</a>
 *      </p>*
 * @author eugeis
 */

public class OutlookBarStorage extends OleAuto {

  protected Variant groups;

  public OutlookBarStorage(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getGroups();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212049(v=office.11).aspx">Groups</a>
   */
  public Variant getGroups() {

    String propertyName = "Groups";
    try {
      if (this.groups == null) {
        this.groups = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.groups;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
