package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211211(v=office.11).aspx">View</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171460(v=office.11).aspx">Language</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171474(v=office.11).aspx">LockUserChanges</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171930(v=office.11).aspx">SaveOption</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172049(v=office.11).aspx">Standard</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221814(v=office.11).aspx">ViewType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221858(v=office.11).aspx">XML</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220072(v=office.11).aspx">Apply</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220080(v=office.11).aspx">Copy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210267(v=office.11).aspx">Reset</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210281(v=office.11).aspx">Save</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210948(v=office.11).aspx">MAPIFolder</a>
 *      </p>*
 * @author eugeis
 */

public class View extends OleAuto {

  protected String language;

  protected Boolean lockUserChanges;

  protected Variant name;

  protected OlViewSaveOptionEnum saveOption;

  protected Boolean standard;

  protected Variant viewType;

  protected String xML;

  public View(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getLanguage();
    getLockUserChanges();
    getName();
    getSaveOption();
    getStandard();
    getViewType();
    getXML();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171460(v=office.11).aspx">Language</a>
   */
  public String getLanguage() {

    String propertyName = "Language";
    try {
      if (this.language == null) {
        this.language = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.language;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171474(v=office.11).aspx">LockUserChanges</a>
   */
  public Boolean getLockUserChanges() {

    String propertyName = "LockUserChanges";
    try {
      if (this.lockUserChanges == null) {
        this.lockUserChanges = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lockUserChanges;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171930(v=office.11).aspx">SaveOption</a>
   */
  public OlViewSaveOptionEnum getSaveOption() {

    String propertyName = "SaveOption";
    try {
      if (this.saveOption == null) {
        this.saveOption = OlViewSaveOptionEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.saveOption;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172049(v=office.11).aspx">Standard</a>
   */
  public Boolean getStandard() {

    String propertyName = "Standard";
    try {
      if (this.standard == null) {
        this.standard = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.standard;
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

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221858(v=office.11).aspx">XML</a>
   */
  public String getXML() {

    String propertyName = "XML";
    try {
      if (this.xML == null) {
        this.xML = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.xML;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
