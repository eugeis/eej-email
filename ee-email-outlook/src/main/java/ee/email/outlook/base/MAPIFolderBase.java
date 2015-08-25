package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.Folders;
import ee.email.outlook.Items;
import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210948(v=office.11).aspx">MAPIFolder</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211350(v=office.11).aspx">AddressBookName</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211822(v=office.11).aspx">CurrentView</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211824(v=office.11).aspx">CustomViewsOnly</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211828(v=office.11).aspx">DefaultItemType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211829(v=office.11).aspx">DefaultMessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211836(v=office.11).aspx">Description</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212014(v=office.11).aspx">FolderPath</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212017(v=office.11).aspx">Folders</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171433(v=office.11).aspx">InAppFolderSyncObject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171450(v=office.11).aspx">IsSharePointFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171455(v=office.11).aspx">Items</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172026(v=office.11).aspx">ShowAsOutlookAB</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172030(v=office.11).aspx">ShowItemCount</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212421(v=office.11).aspx">StoreID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221738(v=office.11).aspx">UnReadItemCount</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221807(v=office.11).aspx">Views</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221842(v=office.11).aspx">WebViewOn</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221847(v=office.11).aspx">WebViewURL</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220068(v=office.11).aspx">AddToFavorites</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220069(v=office.11).aspx">AddToPFFavorites</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220079(v=office.11).aspx">CopyTo</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220101(v=office.11).aspx">GetExplorer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220126(v=office.11).aspx">MoveTo</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210916(v=office.11).aspx">Explorer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210946(v=office.11).aspx">MailItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210951(v=office.11).aspx">MeetingItem</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa211211(v=office.11).aspx">View</a>
 *      </p>*
 * @author eugeis
 */

public class MAPIFolderBase extends OleAuto {

  protected String addressBookName;

  protected Variant currentView;

  protected Boolean customViewsOnly;

  protected OlItemTypeEnum defaultItemType;

  protected String defaultMessageClass;

  protected String description;

  protected String entryID;

  protected String folderPath;

  protected Folders folders;

  protected Boolean inAppFolderSyncObject;

  protected Boolean isSharePointFolder;

  protected Items items;

  protected Variant name;

  protected Variant showAsOutlookAB;

  protected Variant showItemCount;

  protected String storeID;

  protected Variant unReadItemCount;

  protected Views views;

  protected Boolean webViewOn;

  protected String webViewURL;

  public MAPIFolderBase(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getAddressBookName();
    getCurrentView();
    getCustomViewsOnly();
    getDefaultItemType();
    getDefaultMessageClass();
    getDescription();
    getEntryID();
    getFolderPath();
    getFolders();
    getInAppFolderSyncObject();
    getIsSharePointFolder();
    getItems();
    getName();
    getShowAsOutlookAB();
    getShowItemCount();
    getStoreID();
    getUnReadItemCount();
    getViews();
    getWebViewOn();
    getWebViewURL();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211350(v=office.11).aspx">AddressBookName</a>
   */
  public String getAddressBookName() {

    String propertyName = "AddressBookName";
    try {
      if (this.addressBookName == null) {
        this.addressBookName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.addressBookName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211822(v=office.11).aspx">CurrentView</a>
   */
  public Variant getCurrentView() {

    String propertyName = "CurrentView";
    try {
      if (this.currentView == null) {
        this.currentView = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentView;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211824(v=office.11).aspx">CustomViewsOnly</a>
   */
  public Boolean getCustomViewsOnly() {

    String propertyName = "CustomViewsOnly";
    try {
      if (this.customViewsOnly == null) {
        this.customViewsOnly = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.customViewsOnly;
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

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211829(v=office.11).aspx">DefaultMessageClass</a>
   */
  public String getDefaultMessageClass() {

    String propertyName = "DefaultMessageClass";
    try {
      if (this.defaultMessageClass == null) {
        this.defaultMessageClass = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.defaultMessageClass;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211836(v=office.11).aspx">Description</a>
   */
  public String getDescription() {

    String propertyName = "Description";
    try {
      if (this.description == null) {
        this.description = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.description;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a>
   */
  public String getEntryID() {

    String propertyName = "EntryID";
    try {
      if (this.entryID == null) {
        this.entryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.entryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212014(v=office.11).aspx">FolderPath</a>
   */
  public String getFolderPath() {

    String propertyName = "FolderPath";
    try {
      if (this.folderPath == null) {
        this.folderPath = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.folderPath;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171433(v=office.11).aspx">InAppFolderSyncObject</a>
   */
  public Boolean getInAppFolderSyncObject() {

    String propertyName = "InAppFolderSyncObject";
    try {
      if (this.inAppFolderSyncObject == null) {
        this.inAppFolderSyncObject = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.inAppFolderSyncObject;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171450(v=office.11).aspx">IsSharePointFolder</a>
   */
  public Boolean getIsSharePointFolder() {

    String propertyName = "IsSharePointFolder";
    try {
      if (this.isSharePointFolder == null) {
        this.isSharePointFolder = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isSharePointFolder;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171455(v=office.11).aspx">Items</a>
   */
  public Items getItems() {

    String propertyName = "Items";
    try {
      if (this.items == null) {
        this.items = new Items(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.items;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172026(v=office.11).aspx">ShowAsOutlookAB</a>
   */
  public Variant getShowAsOutlookAB() {

    String propertyName = "ShowAsOutlookAB";
    try {
      if (this.showAsOutlookAB == null) {
        this.showAsOutlookAB = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.showAsOutlookAB;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172030(v=office.11).aspx">ShowItemCount</a>
   */
  public Variant getShowItemCount() {

    String propertyName = "ShowItemCount";
    try {
      if (this.showItemCount == null) {
        this.showItemCount = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.showItemCount;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212421(v=office.11).aspx">StoreID</a>
   */
  public String getStoreID() {

    String propertyName = "StoreID";
    try {
      if (this.storeID == null) {
        this.storeID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.storeID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221738(v=office.11).aspx">UnReadItemCount</a>
   */
  public Variant getUnReadItemCount() {

    String propertyName = "UnReadItemCount";
    try {
      if (this.unReadItemCount == null) {
        this.unReadItemCount = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.unReadItemCount;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221807(v=office.11).aspx">Views</a>
   */
  public Views getViews() {

    String propertyName = "Views";
    try {
      if (this.views == null) {
        this.views = new Views(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.views;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221842(v=office.11).aspx">WebViewOn</a>
   */
  public Boolean getWebViewOn() {

    String propertyName = "WebViewOn";
    try {
      if (this.webViewOn == null) {
        this.webViewOn = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.webViewOn;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221847(v=office.11).aspx">WebViewURL</a>
   */
  public String getWebViewURL() {

    String propertyName = "WebViewURL";
    try {
      if (this.webViewURL == null) {
        this.webViewURL = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.webViewURL;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.folders != null) {
      this.folders.dispose();
    }
    if (this.items != null) {
      this.items.dispose();
    }
    if (this.views != null) {
      this.views.dispose();
    }
  }

}
