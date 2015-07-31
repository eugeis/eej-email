package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210887(v=office.11).aspx">Action</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211815(v=office.11).aspx">CopyLike</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211861(v=office.11).aspx">Enabled</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171847(v=office.11).aspx">Prefix</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171912(v=office.11).aspx">ReplyStyle</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171924(v=office.11).aspx">ResponseStyle</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172038(v=office.11).aspx">ShowOn</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220091(v=office.11).aspx">Execute</a>
 *      </p>*
 * @author eugeis
 */

public class Action extends OleAuto {

  protected OlActionCopyLikeEnum copyLike;

  protected Boolean enabled;

  protected String messageClass;

  protected Variant name;

  protected String prefix;

  protected OlActionReplyStyleEnum replyStyle;

  protected OlActionResponseStyleEnum responseStyle;

  protected OlActionShowOnEnum showOn;

  public Action(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getCopyLike();
    getEnabled();
    getMessageClass();
    getName();
    getPrefix();
    getReplyStyle();
    getResponseStyle();
    getShowOn();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211815(v=office.11).aspx">CopyLike</a>
   */
  public OlActionCopyLikeEnum getCopyLike() {

    String propertyName = "CopyLike";
    try {
      if (this.copyLike == null) {
        this.copyLike = OlActionCopyLikeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.copyLike;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211861(v=office.11).aspx">Enabled</a>
   */
  public Boolean getEnabled() {

    String propertyName = "Enabled";
    try {
      if (this.enabled == null) {
        this.enabled = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.enabled;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171847(v=office.11).aspx">Prefix</a>
   */
  public String getPrefix() {

    String propertyName = "Prefix";
    try {
      if (this.prefix == null) {
        this.prefix = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.prefix;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171912(v=office.11).aspx">ReplyStyle</a>
   */
  public OlActionReplyStyleEnum getReplyStyle() {

    String propertyName = "ReplyStyle";
    try {
      if (this.replyStyle == null) {
        this.replyStyle = OlActionReplyStyleEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.replyStyle;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171924(v=office.11).aspx">ResponseStyle</a>
   */
  public OlActionResponseStyleEnum getResponseStyle() {

    String propertyName = "ResponseStyle";
    try {
      if (this.responseStyle == null) {
        this.responseStyle = OlActionResponseStyleEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.responseStyle;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172038(v=office.11).aspx">ShowOn</a>
   */
  public OlActionShowOnEnum getShowOn() {

    String propertyName = "ShowOn";
    try {
      if (this.showOn == null) {
        this.showOn = OlActionShowOnEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.showOn;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
