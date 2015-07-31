package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.DefaultOleAutoFactory;
import ee.email.outlook.OleAuto;
import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210918(v=office.11).aspx">Folders</a>
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
 *      href="http://msdn.microsoft.com/en-us/library/aa220102(v=office.11).aspx">GetFirst</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220106(v=office.11).aspx">GetLast</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220109(v=office.11).aspx">GetNext</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220113(v=office.11).aspx">GetPrevious</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171238(v=office.11).aspx">FolderAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171245(v=office.11).aspx">FolderChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171250(v=office.11).aspx">FolderRemove</a>
 *      </p>*
 * @author eugeis
 */

public class FoldersBase<E extends OleAuto> extends OleCollection<E> {

  public FoldersBase(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new DefaultOleAutoFactory());
  }

  public FoldersBase(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

    super(auto, initImmediate, childFactory);
  }

  @Override
  public void init() {

    super.init();
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
