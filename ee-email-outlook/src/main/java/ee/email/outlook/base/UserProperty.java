package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211200(v=office.11).aspx">UserProperty</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212026(v=office.11).aspx">Formula</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221782(v=office.11).aspx">ValidationFormula</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221789(v=office.11).aspx">ValidationText</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221795(v=office.11).aspx">Value</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a>
 *      </p>*
 * @author eugeis
 */

public class UserProperty extends OleAuto {

  protected String formula;

  protected Variant name;

  protected Variant type;

  protected String validationFormula;

  protected String validationText;

  protected Variant value;

  public UserProperty(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getFormula();
    getName();
    getType();
    getValidationFormula();
    getValidationText();
    getValue();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212026(v=office.11).aspx">Formula</a>
   */
  public String getFormula() {

    String propertyName = "Formula";
    try {
      if (this.formula == null) {
        this.formula = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.formula;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221782(v=office.11).aspx">ValidationFormula</a>
   */
  public String getValidationFormula() {

    String propertyName = "ValidationFormula";
    try {
      if (this.validationFormula == null) {
        this.validationFormula = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.validationFormula;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221789(v=office.11).aspx">ValidationText</a>
   */
  public String getValidationText() {

    String propertyName = "ValidationText";
    try {
      if (this.validationText == null) {
        this.validationText = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.validationText;
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
