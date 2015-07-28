package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.base.FoldersBase;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210918(v=office.11).aspx">Folders</a>
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
 *      </p>
 *      *
 * @author eugeis
 */
public class Folders<E extends OleAuto> extends FoldersBase<E> {

  public Folders(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public Folders(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

    super(auto, initImmediate, childFactory);
  }
}
