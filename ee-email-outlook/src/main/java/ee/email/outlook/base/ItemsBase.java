package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210932(v=office.11).aspx">Items</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171434(v=office.11).aspx">IncludeRecurrences</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220070(v=office.11).aspx">Add</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220093(v=office.11).aspx">Find</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220092(v=office.11).aspx">FindNext</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220102(v=office.11).aspx">GetFirst</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220106(v=office.11).aspx">GetLast</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220109(v=office.11).aspx">GetNext</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220113(v=office.11).aspx">GetPrevious</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210266(v=office.11).aspx">ResetColumns</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210275(v=office.11).aspx">Restrict</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210287(v=office.11).aspx">SetColumns</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210304(v=office.11).aspx">Sort</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171270(v=office.11).aspx">ItemAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171274(v=office.11).aspx">ItemChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171280(v=office.11).aspx">ItemRemove</a>
 *      </p>*
 * @author eugeis
 */

public class ItemsBase extends OleAuto {

  protected Variant count;

  protected Boolean includeRecurrences;

  public ItemsBase(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getCount();
    getIncludeRecurrences();
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

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171434(v=office.11).aspx">IncludeRecurrences</a>
   */
  public Boolean getIncludeRecurrences() {

    String propertyName = "IncludeRecurrences";
    try {
      if (this.includeRecurrences == null) {
        this.includeRecurrences = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.includeRecurrences;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
