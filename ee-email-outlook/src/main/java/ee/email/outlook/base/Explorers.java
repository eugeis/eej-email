package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210914(v=office.11).aspx">Explorers</a>
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
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171297(v=office.11).aspx">NewExplorer</a>
 *      </p>*
 * @author eugeis
 */

public class Explorers<E extends Explorer> extends OleCollection<E> {

  public Explorers(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new OleAutoFactory<E>() {

      @SuppressWarnings("unchecked")
      @Override
      public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

        return (E) new Explorer(auto, initImmediate);
      }
    });

  }

  public Explorers(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

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
