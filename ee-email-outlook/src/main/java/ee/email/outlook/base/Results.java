package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.DefaultOleAutoFactory;
import ee.email.outlook.OleAuto;
import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211043(v=office.11).aspx">Results</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211828(v=office.11).aspx">DefaultItemType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220102(v=office.11).aspx">GetFirst</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220106(v=office.11).aspx">GetLast</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220109(v=office.11).aspx">GetNext</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220113(v=office.11).aspx">GetPrevious</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210266(v=office.11).aspx">ResetColumns</a> | <a
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

public class Results<E extends OleAuto> extends OleCollection<E> {

  protected OlItemTypeEnum defaultItemType;

  public Results(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new DefaultOleAutoFactory());
  }

  public Results(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

    super(auto, initImmediate, childFactory);
  }

  public void init() {

    super.init();
    getDefaultItemType();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211828(v=office.11).aspx">DefaultItemType</a>
   */
  public OlItemTypeEnum getDefaultItemType() {

    String propertyName = "DefaultItemType";
    try {
      if (this.defaultItemType == null) {
        this.defaultItemType = OlItemTypeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.defaultItemType;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
