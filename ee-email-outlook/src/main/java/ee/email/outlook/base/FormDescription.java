package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210920(v=office.11).aspx">FormDescription</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211790(v=office.11).aspx">Category</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211789(v=office.11).aspx">CategorySub</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211797(v=office.11).aspx">Comment</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211811(v=office.11).aspx">ContactName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211838(v=office.11).aspx">DisplayName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171371(v=office.11).aspx">Hidden</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171428(v=office.11).aspx">Icon</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171473(v=office.11).aspx">Locked</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171493(v=office.11).aspx">MiniIcon</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171771(v=office.11).aspx">Number</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171781(v=office.11).aspx">OneOff</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171820(v=office.11).aspx">Password</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171936(v=office.11).aspx">ScriptText</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220297(v=office.11).aspx">Template</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221779(v=office.11).aspx">UseWordMail</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221802(v=office.11).aspx">Version</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220132(v=office.11).aspx">PublishForm</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210899(v=office.11).aspx">AppointmentItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210907(v=office.11).aspx">ContactItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210909(v=office.11).aspx">DistListItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210910(v=office.11).aspx">DocumentItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210936(v=office.11).aspx">JournalItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210946(v=office.11).aspx">MailItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210951(v=office.11).aspx">MeetingItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210983(v=office.11).aspx">PostItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211033(v=office.11).aspx">RemoteItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211038(v=office.11).aspx">ReportItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211067(v=office.11).aspx">TaskItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211069(v=office.11).aspx">TaskRequestAcceptItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211079(v=office.11).aspx">TaskRequestDeclineItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211086(v=office.11).aspx">TaskRequestItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211090(v=office.11).aspx">TaskRequestUpdateItem</a>
 *      </p>*
 * @author eugeis
 */

public class FormDescription extends OleAuto {

  protected String category;

  protected String categorySub;

  protected String comment;

  protected String contactName;

  protected String displayName;

  protected Boolean hidden;

  protected String icon;

  protected Boolean locked;

  protected String messageClass;

  protected String miniIcon;

  protected Variant name;

  protected String number;

  protected Boolean oneOff;

  protected String password;

  protected String scriptText;

  protected String template;

  protected Boolean useWordMail;

  protected String version;

  public FormDescription(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getCategory();
    getCategorySub();
    getComment();
    getContactName();
    getDisplayName();
    getHidden();
    getIcon();
    getLocked();
    getMessageClass();
    getMiniIcon();
    getName();
    getNumber();
    getOneOff();
    getPassword();
    getScriptText();
    getTemplate();
    getUseWordMail();
    getVersion();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211790(v=office.11).aspx">Category</a>
   */
  public String getCategory() {

    String propertyName = "Category";
    try {
      if (this.category == null) {
        this.category = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.category;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211789(v=office.11).aspx">CategorySub</a>
   */
  public String getCategorySub() {

    String propertyName = "CategorySub";
    try {
      if (this.categorySub == null) {
        this.categorySub = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.categorySub;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211797(v=office.11).aspx">Comment</a>
   */
  public String getComment() {

    String propertyName = "Comment";
    try {
      if (this.comment == null) {
        this.comment = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.comment;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211811(v=office.11).aspx">ContactName</a>
   */
  public String getContactName() {

    String propertyName = "ContactName";
    try {
      if (this.contactName == null) {
        this.contactName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.contactName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211838(v=office.11).aspx">DisplayName</a>
   */
  public String getDisplayName() {

    String propertyName = "DisplayName";
    try {
      if (this.displayName == null) {
        this.displayName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.displayName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171371(v=office.11).aspx">Hidden</a>
   */
  public Boolean getHidden() {

    String propertyName = "Hidden";
    try {
      if (this.hidden == null) {
        this.hidden = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hidden;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171428(v=office.11).aspx">Icon</a>
   */
  public String getIcon() {

    String propertyName = "Icon";
    try {
      if (this.icon == null) {
        this.icon = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.icon;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171473(v=office.11).aspx">Locked</a>
   */
  public Boolean getLocked() {

    String propertyName = "Locked";
    try {
      if (this.locked == null) {
        this.locked = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.locked;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a>
   */
  public String getMessageClass() {

    String propertyName = "MessageClass";
    try {
      if (this.messageClass == null) {
        this.messageClass = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.messageClass;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171493(v=office.11).aspx">MiniIcon</a>
   */
  public String getMiniIcon() {

    String propertyName = "MiniIcon";
    try {
      if (this.miniIcon == null) {
        this.miniIcon = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.miniIcon;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171771(v=office.11).aspx">Number</a>
   */
  public String getNumber() {

    String propertyName = "Number";
    try {
      if (this.number == null) {
        this.number = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.number;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171781(v=office.11).aspx">OneOff</a>
   */
  public Boolean getOneOff() {

    String propertyName = "OneOff";
    try {
      if (this.oneOff == null) {
        this.oneOff = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.oneOff;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171820(v=office.11).aspx">Password</a>
   */
  public String getPassword() {

    String propertyName = "Password";
    try {
      if (this.password == null) {
        this.password = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.password;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171936(v=office.11).aspx">ScriptText</a>
   */
  public String getScriptText() {

    String propertyName = "ScriptText";
    try {
      if (this.scriptText == null) {
        this.scriptText = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.scriptText;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220297(v=office.11).aspx">Template</a>
   */
  public String getTemplate() {

    String propertyName = "Template";
    try {
      if (this.template == null) {
        this.template = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.template;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221779(v=office.11).aspx">UseWordMail</a>
   */
  public Boolean getUseWordMail() {

    String propertyName = "UseWordMail";
    try {
      if (this.useWordMail == null) {
        this.useWordMail = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.useWordMail;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221802(v=office.11).aspx">Version</a>
   */
  public String getVersion() {

    String propertyName = "Version";
    try {
      if (this.version == null) {
        this.version = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.version;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
