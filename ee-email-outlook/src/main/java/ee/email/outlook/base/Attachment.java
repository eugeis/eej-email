package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210902(v=office.11).aspx">Attachment</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211838(v=office.11).aspx">DisplayName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211870(v=office.11).aspx">FileName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171435(v=office.11).aspx">Index</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171824(v=office.11).aspx">PathName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171845(v=office.11).aspx">Position</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210277(v=office.11).aspx">SaveAsFile</a>
 *      </p>*
 * @author eugeis
 */

public class Attachment extends OleAuto {

  protected String displayName;

  protected String fileName;

  protected Variant index;

  protected String pathName;

  protected Variant position;

  protected Variant type;

  public Attachment(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getDisplayName();
    getFileName();
    getIndex();
    getPathName();
    getPosition();
    getType();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211838(v=office.11).aspx">DisplayName</a>
   */
  public String getDisplayName() {

    String propertyName = "DisplayName";
    try {
      if (this.displayName == null) {
        this.displayName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.displayName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211870(v=office.11).aspx">FileName</a>
   */
  public String getFileName() {

    String propertyName = "FileName";
    try {
      if (this.fileName == null) {
        this.fileName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.fileName;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171824(v=office.11).aspx">PathName</a>
   */
  public String getPathName() {

    String propertyName = "PathName";
    try {
      if (this.pathName == null) {
        this.pathName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.pathName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171845(v=office.11).aspx">Position</a>
   */
  public Variant getPosition() {

    String propertyName = "Position";
    try {
      if (this.position == null) {
        this.position = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.position;
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
