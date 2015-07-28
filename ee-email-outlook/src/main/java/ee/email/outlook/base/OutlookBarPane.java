package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210966(v=office.11).aspx">OutlookBarPane</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211812(v=office.11).aspx">Contents</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211819(v=office.11).aspx">CurrentGroup</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221819(v=office.11).aspx">Visible</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171108(v=office.11).aspx">BeforeGroupSwitch</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171184(v=office.11).aspx">BeforeNavigate</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210962(v=office.11).aspx">OutlookBarGroup</a>
 *      | <a href="http://msdn.microsoft.com/en-us/library/aa210976(v=office.11).aspx">OutlookBarStorage</a>
 *      </p>*
 * @author eugeis
 */

public class OutlookBarPane extends OleAuto {

  protected Variant contents;

  protected Variant currentGroup;

  protected Variant name;

  protected Boolean visible;

  public OutlookBarPane(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getContents();
    getCurrentGroup();
    getName();
    getVisible();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211812(v=office.11).aspx">Contents</a>
   */
  public Variant getContents() {

    String propertyName = "Contents";
    try {
      if (this.contents == null) {
        this.contents = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.contents;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211819(v=office.11).aspx">CurrentGroup</a>
   */
  public Variant getCurrentGroup() {

    String propertyName = "CurrentGroup";
    try {
      if (this.currentGroup == null) {
        this.currentGroup = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentGroup;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221819(v=office.11).aspx">Visible</a>
   */
  public Boolean getVisible() {

    String propertyName = "Visible";
    try {
      if (this.visible == null) {
        this.visible = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.visible;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
