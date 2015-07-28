package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.base.NameSpaceBase;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210954(v=office.11).aspx">NameSpace</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa210225(v=office.11).aspx">AddStore</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219415(v=office.11).aspx">AddStoreEx</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220084(v=office.11).aspx">CreateRecipient</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220088(v=office.11).aspx">Dial</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220100(v=office.11).aspx">GetDefaultFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220103(v=office.11).aspx">GetFolderFromID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220105(v=office.11).aspx">GetItemFromID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220114(v=office.11).aspx">GetRecipientFromID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220116(v=office.11).aspx">GetSharedDefaultFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220123(v=office.11).aspx">Logoff</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220124(v=office.11).aspx">Logon</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220129(v=office.11).aspx">PickFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220140(v=office.11).aspx">RemoveStore</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171319(v=office.11).aspx">OptionsPagesAdd</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210893(v=office.11).aspx">AddressLists</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa211006(v=office.11).aspx">Recipient</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211054(v=office.11).aspx">SyncObjects</a>
 *      </p>
 *      *
 * @author eugeis
 */
public class NameSpace extends NameSpaceBase {

  public NameSpace(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

}
