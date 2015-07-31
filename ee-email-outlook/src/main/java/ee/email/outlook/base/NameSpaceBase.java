package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.Folders;
import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210954(v=office.11).aspx">NameSpace</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211361(v=office.11).aspx">AddressLists</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211821(v=office.11).aspx">CurrentUser</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211809(v=office.11).aspx">ExchangeConnectionMode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212017(v=office.11).aspx">Folders</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171776(v=office.11).aspx">Offline</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220160(v=office.11).aspx">SyncObjects</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220737(v=office.11).aspx">Type</a>
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
 *      </p>*
 * @author eugeis
 */

public class NameSpaceBase extends OleAuto {

  protected AddressLists addressLists;

  protected Variant currentUser;

  protected Variant exchangeConnectionMode;

  protected Folders folders;

  protected Boolean offline;

  protected SyncObjects syncObjects;

  protected Variant type;

  public NameSpaceBase(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getAddressLists();
    getCurrentUser();
    getExchangeConnectionMode();
    getFolders();
    getOffline();
    getSyncObjects();
    getType();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211361(v=office.11).aspx">AddressLists</a>
   */
  public AddressLists getAddressLists() {

    String propertyName = "AddressLists";
    try {
      if (this.addressLists == null) {
        this.addressLists = new AddressLists(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.addressLists;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211821(v=office.11).aspx">CurrentUser</a>
   */
  public Variant getCurrentUser() {

    String propertyName = "CurrentUser";
    try {
      if (this.currentUser == null) {
        this.currentUser = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentUser;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211809(v=office.11).aspx">ExchangeConnectionMode</a>
   */
  public Variant getExchangeConnectionMode() {

    String propertyName = "ExchangeConnectionMode";
    try {
      if (this.exchangeConnectionMode == null) {
        this.exchangeConnectionMode = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.exchangeConnectionMode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212017(v=office.11).aspx">Folders</a>
   */
  public Folders getFolders() {

    String propertyName = "Folders";
    try {
      if (this.folders == null) {
        this.folders = new Folders(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.folders;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171776(v=office.11).aspx">Offline</a>
   */
  public Boolean getOffline() {

    String propertyName = "Offline";
    try {
      if (this.offline == null) {
        this.offline = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.offline;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220160(v=office.11).aspx">SyncObjects</a>
   */
  public SyncObjects getSyncObjects() {

    String propertyName = "SyncObjects";
    try {
      if (this.syncObjects == null) {
        this.syncObjects = new SyncObjects(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.syncObjects;
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
    if (this.addressLists != null) {
      this.addressLists.dispose();
    }
    if (this.folders != null) {
      this.folders.dispose();
    }
    if (this.syncObjects != null) {
      this.syncObjects.dispose();
    }
  }

}
