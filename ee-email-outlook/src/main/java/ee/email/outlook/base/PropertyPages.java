package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210986(v=office.11).aspx">PropertyPages</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220070(v=office.11).aspx">Add</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a>
 *      </p>*
 * @author eugeis
 */

public class PropertyPages<E extends PropertyPage> extends OleCollection<E> {

  public PropertyPages(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new OleAutoFactory<E>() {

      @SuppressWarnings("unchecked")
      @Override
      public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

        return (E) new PropertyPage(auto, initImmediate);
      }
    });

  }

  public PropertyPages(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

    super(auto, initImmediate, childFactory);
  }

  public void init() {

    super.init();
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
