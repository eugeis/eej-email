package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210946(v=office.11).aspx">MailItem</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211372(v=office.11).aspx">AlternateRecipientAllowed</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211411(v=office.11).aspx">AutoForwarded</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211422(v=office.11).aspx">BCC</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211430(v=office.11).aspx">BodyFormat</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211791(v=office.11).aspx">CC</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211830(v=office.11).aspx">DeferredDeliveryTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211833(v=office.11).aspx">DeleteAfterSubmit</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211862(v=office.11).aspx">EnableSharedAttachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211867(v=office.11).aspx">ExpiryTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211979(v=office.11).aspx">FlagDueBy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211991(v=office.11).aspx">FlagIcon</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212007(v=office.11).aspx">FlagRequest</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212013(v=office.11).aspx">FlagStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171418(v=office.11).aspx">HTMLBody</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212147(v=office.11).aspx">HasCoverSheet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171439(v=office.11).aspx">InternetCodepage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171446(v=office.11).aspx">IsIPFax</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171788(v=office.11).aspx">OriginatorDeliveryReportRequested</a>
 *      | <a href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171840(v=office.11).aspx">Permission</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171838(v=office.11).aspx">PermissionService</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171860(v=office.11).aspx">ReadReceiptRequested</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171862(v=office.11).aspx">ReceivedByEntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171864(v=office.11).aspx">ReceivedByName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171866(v=office.11).aspx">ReceivedOnBehalfOfEntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171869(v=office.11).aspx">ReceivedOnBehalfOfName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171873(v=office.11).aspx">ReceivedTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171876(v=office.11).aspx">RecipientReassignmentProhibited</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa171878(v=office.11).aspx">Recipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171894(v=office.11).aspx">ReminderOverrideDefault</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171897(v=office.11).aspx">ReminderPlaySound</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171899(v=office.11).aspx">ReminderSet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171901(v=office.11).aspx">ReminderSoundFile</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171904(v=office.11).aspx">ReminderTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171908(v=office.11).aspx">RemoteStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171909(v=office.11).aspx">ReplyRecipientNames</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171911(v=office.11).aspx">ReplyRecipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171932(v=office.11).aspx">SaveSentMessageFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171942(v=office.11).aspx">SenderEmailAddress</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171944(v=office.11).aspx">SenderEmailType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171985(v=office.11).aspx">SenderName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172010(v=office.11).aspx">Sent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172005(v=office.11).aspx">SentOn</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171998(v=office.11).aspx">SentOnBehalfOfName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220151(v=office.11).aspx">Submitted</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220513(v=office.11).aspx">To</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221825(v=office.11).aspx">VotingOptions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221831(v=office.11).aspx">VotingResponse</a>
 *      </p>
 *      <p>
 *      Methods | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220075(v=office.11).aspx">ClearConversationIndex</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220080(v=office.11).aspx">Copy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220096(v=office.11).aspx">Forward</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220127(v=office.11).aspx">Move</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220131(v=office.11).aspx">PrintOut</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220146(v=office.11).aspx">Reply</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220145(v=office.11).aspx">ReplyAll</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210281(v=office.11).aspx">Save</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210279(v=office.11).aspx">SaveAs</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210285(v=office.11).aspx">Send</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210294(v=office.11).aspx">ShowCategoriesDialog</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa209975(v=office.11).aspx">AttachmentAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209976(v=office.11).aspx">AttachmentRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209977(v=office.11).aspx">BeforeAttachmentSave</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209978(v=office.11).aspx">BeforeCheckNames</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209979(v=office.11).aspx">BeforeDelete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171213(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171218(v=office.11).aspx">CustomAction</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171227(v=office.11).aspx">CustomPropertyChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171259(v=office.11).aspx">Forward</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171315(v=office.11).aspx">Open</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171326(v=office.11).aspx">PropertyChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171358(v=office.11).aspx">Read</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171365(v=office.11).aspx">Reply</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171364(v=office.11).aspx">ReplyAll</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219360(v=office.11).aspx">Send</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219369(v=office.11).aspx">Write</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210886(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210901(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210904(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210920(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210924(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210939(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210948(v=office.11).aspx">MAPIFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210995(v=office.11).aspx">Recipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211095(v=office.11).aspx">UserProperties</a>
 *      </p>*
 * @author eugeis
 */

public class MailItem extends OleAuto {

  protected Actions actions;

  protected Boolean alternateRecipientAllowed;

  protected Attachments attachments;

  protected Boolean autoForwarded;

  protected Boolean autoResolvedWinner;

  protected String bCC;

  protected String billingInformation;

  protected String body;

  protected OlBodyFormatEnum bodyFormat;

  protected String cC;

  protected String categories;

  protected String companies;

  protected Conflicts conflicts;

  protected String conversationIndex;

  protected String conversationTopic;

  protected Date creationTime;

  protected Date deferredDeliveryTime;

  protected Boolean deleteAfterSubmit;

  protected OlDownloadStateEnum downloadState;

  protected Variant enableSharedAttachments;

  protected String entryID;

  protected Date expiryTime;

  protected Date flagDueBy;

  protected Variant flagIcon;

  protected String flagRequest;

  protected OlFlagStatusEnum flagStatus;

  protected FormDescription formDescription;

  protected Variant getInspector;

  protected String hTMLBody;

  protected Variant hasCoverSheet;

  protected OlImportanceEnum importance;

  protected Variant internetCodepage;

  protected Boolean isConflict;

  protected Variant isIPFax;

  protected ItemProperties itemProperties;

  protected Date lastModificationTime;

  protected Links links;

  protected OlRemoteStatusEnum markForDownload;

  protected String messageClass;

  protected String mileage;

  protected Boolean noAging;

  protected Boolean originatorDeliveryReportRequested;

  protected Variant outlookInternalVersion;

  protected String outlookVersion;

  protected Variant permission;

  protected Variant permissionService;

  protected Variant readReceiptRequested;

  protected String receivedByEntryID;

  protected String receivedByName;

  protected String receivedOnBehalfOfEntryID;

  protected String receivedOnBehalfOfName;

  protected Date receivedTime;

  protected Boolean recipientReassignmentProhibited;

  protected Recipients recipients;

  protected Boolean reminderOverrideDefault;

  protected Boolean reminderPlaySound;

  protected Boolean reminderSet;

  protected String reminderSoundFile;

  protected Date reminderTime;

  protected OlRemoteStatusEnum remoteStatus;

  protected String replyRecipientNames;

  protected Variant replyRecipients;

  protected Variant saveSentMessageFolder;

  protected Boolean saved;

  protected String senderEmailAddress;

  protected String senderEmailType;

  protected String senderName;

  protected OlSensitivityEnum sensitivity;

  protected Boolean sent;

  protected Date sentOn;

  protected String sentOnBehalfOfName;

  protected Variant size;

  protected String subject;

  protected Boolean submitted;

  protected String to;

  protected Boolean unRead;

  protected UserProperties userProperties;

  protected String votingOptions;

  protected String votingResponse;

  public MailItem(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getActions();
    getAlternateRecipientAllowed();
    getAttachments();
    getAutoForwarded();
    getAutoResolvedWinner();
    getBCC();
    getBillingInformation();
    getBody();
    getBodyFormat();
    getCC();
    getCategories();
    getCompanies();
    getConflicts();
    getConversationIndex();
    getConversationTopic();
    getCreationTime();
    getDeferredDeliveryTime();
    getDeleteAfterSubmit();
    getDownloadState();
    getEnableSharedAttachments();
    getEntryID();
    getExpiryTime();
    getFlagDueBy();
    getFlagIcon();
    getFlagRequest();
    getFlagStatus();
    getFormDescription();
    getGetInspector();
    getHTMLBody();
    getHasCoverSheet();
    getImportance();
    getInternetCodepage();
    getIsConflict();
    getIsIPFax();
    getItemProperties();
    getLastModificationTime();
    getLinks();
    getMarkForDownload();
    getMessageClass();
    getMileage();
    getNoAging();
    getOriginatorDeliveryReportRequested();
    getOutlookInternalVersion();
    getOutlookVersion();
    getPermission();
    getPermissionService();
    getReadReceiptRequested();
    getReceivedByEntryID();
    getReceivedByName();
    getReceivedOnBehalfOfEntryID();
    getReceivedOnBehalfOfName();
    getReceivedTime();
    getRecipientReassignmentProhibited();
    getRecipients();
    getReminderOverrideDefault();
    getReminderPlaySound();
    getReminderSet();
    getReminderSoundFile();
    getReminderTime();
    getRemoteStatus();
    getReplyRecipientNames();
    getReplyRecipients();
    getSaveSentMessageFolder();
    getSaved();
    getSenderEmailAddress();
    getSenderEmailType();
    getSenderName();
    getSensitivity();
    getSent();
    getSentOn();
    getSentOnBehalfOfName();
    getSize();
    getSubject();
    getSubmitted();
    getTo();
    getUnRead();
    getUserProperties();
    getVotingOptions();
    getVotingResponse();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a>
   */
  public Actions getActions() {

    String propertyName = "Actions";
    try {
      if (this.actions == null) {
        this.actions = new Actions(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.actions;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211372(v=office.11).aspx">AlternateRecipientAllowed</a>
   */
  public Boolean getAlternateRecipientAllowed() {

    String propertyName = "AlternateRecipientAllowed";
    try {
      if (this.alternateRecipientAllowed == null) {
        this.alternateRecipientAllowed = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.alternateRecipientAllowed;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a>
   */
  public Attachments getAttachments() {

    String propertyName = "Attachments";
    try {
      if (this.attachments == null) {
        this.attachments = new Attachments(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.attachments;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211411(v=office.11).aspx">AutoForwarded</a>
   */
  public Boolean getAutoForwarded() {

    String propertyName = "AutoForwarded";
    try {
      if (this.autoForwarded == null) {
        this.autoForwarded = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.autoForwarded;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a>
   */
  public Boolean getAutoResolvedWinner() {

    String propertyName = "AutoResolvedWinner";
    try {
      if (this.autoResolvedWinner == null) {
        this.autoResolvedWinner = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.autoResolvedWinner;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211422(v=office.11).aspx">BCC</a>
   */
  public String getBCC() {

    String propertyName = "BCC";
    try {
      if (this.bCC == null) {
        this.bCC = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.bCC;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a>
   */
  public String getBillingInformation() {

    String propertyName = "BillingInformation";
    try {
      if (this.billingInformation == null) {
        this.billingInformation = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.billingInformation;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a>
   */
  public String getBody() {

    String propertyName = "Body";
    try {
      if (this.body == null) {
        this.body = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.body;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211430(v=office.11).aspx">BodyFormat</a>
   */
  public OlBodyFormatEnum getBodyFormat() {

    String propertyName = "BodyFormat";
    try {
      if (this.bodyFormat == null) {
        this.bodyFormat = OlBodyFormatEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.bodyFormat;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211791(v=office.11).aspx">CC</a>
   */
  public String getCC() {

    String propertyName = "CC";
    try {
      if (this.cC == null) {
        this.cC = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.cC;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a>
   */
  public String getCategories() {

    String propertyName = "Categories";
    try {
      if (this.categories == null) {
        this.categories = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.categories;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a>
   */
  public String getCompanies() {

    String propertyName = "Companies";
    try {
      if (this.companies == null) {
        this.companies = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.companies;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a>
   */
  public Conflicts getConflicts() {

    String propertyName = "Conflicts";
    try {
      if (this.conflicts == null) {
        this.conflicts = new Conflicts(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conflicts;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a>
   */
  public String getConversationIndex() {

    String propertyName = "ConversationIndex";
    try {
      if (this.conversationIndex == null) {
        this.conversationIndex = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conversationIndex;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a>
   */
  public String getConversationTopic() {

    String propertyName = "ConversationTopic";
    try {
      if (this.conversationTopic == null) {
        this.conversationTopic = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conversationTopic;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a>
   */
  public Date getCreationTime() {

    String propertyName = "CreationTime";
    try {
      if (this.creationTime == null) {
        this.creationTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.creationTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211830(v=office.11).aspx">DeferredDeliveryTime</a>
   */
  public Date getDeferredDeliveryTime() {

    String propertyName = "DeferredDeliveryTime";
    try {
      if (this.deferredDeliveryTime == null) {
        this.deferredDeliveryTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.deferredDeliveryTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211833(v=office.11).aspx">DeleteAfterSubmit</a>
   */
  public Boolean getDeleteAfterSubmit() {

    String propertyName = "DeleteAfterSubmit";
    try {
      if (this.deleteAfterSubmit == null) {
        this.deleteAfterSubmit = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.deleteAfterSubmit;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a>
   */
  public OlDownloadStateEnum getDownloadState() {

    String propertyName = "DownloadState";
    try {
      if (this.downloadState == null) {
        this.downloadState = OlDownloadStateEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.downloadState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211862(v=office.11).aspx">EnableSharedAttachments</a>
   */
  public Variant getEnableSharedAttachments() {

    String propertyName = "EnableSharedAttachments";
    try {
      if (this.enableSharedAttachments == null) {
        this.enableSharedAttachments = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.enableSharedAttachments;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211867(v=office.11).aspx">ExpiryTime</a>
   */
  public Date getExpiryTime() {

    String propertyName = "ExpiryTime";
    try {
      if (this.expiryTime == null) {
        this.expiryTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.expiryTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211979(v=office.11).aspx">FlagDueBy</a>
   */
  public Date getFlagDueBy() {

    String propertyName = "FlagDueBy";
    try {
      if (this.flagDueBy == null) {
        this.flagDueBy = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.flagDueBy;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211991(v=office.11).aspx">FlagIcon</a>
   */
  public Variant getFlagIcon() {

    String propertyName = "FlagIcon";
    try {
      if (this.flagIcon == null) {
        this.flagIcon = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.flagIcon;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212007(v=office.11).aspx">FlagRequest</a>
   */
  public String getFlagRequest() {

    String propertyName = "FlagRequest";
    try {
      if (this.flagRequest == null) {
        this.flagRequest = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.flagRequest;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212013(v=office.11).aspx">FlagStatus</a>
   */
  public OlFlagStatusEnum getFlagStatus() {

    String propertyName = "FlagStatus";
    try {
      if (this.flagStatus == null) {
        this.flagStatus = OlFlagStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.flagStatus;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a>
   */
  public FormDescription getFormDescription() {

    String propertyName = "FormDescription";
    try {
      if (this.formDescription == null) {
        this.formDescription = new FormDescription(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.formDescription;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a>
   */
  public Variant getGetInspector() {

    String propertyName = "GetInspector";
    try {
      if (this.getInspector == null) {
        this.getInspector = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.getInspector;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171418(v=office.11).aspx">HTMLBody</a>
   */
  public String getHTMLBody() {

    String propertyName = "HTMLBody";
    try {
      if (this.hTMLBody == null) {
        this.hTMLBody = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hTMLBody;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212147(v=office.11).aspx">HasCoverSheet</a>
   */
  public Variant getHasCoverSheet() {

    String propertyName = "HasCoverSheet";
    try {
      if (this.hasCoverSheet == null) {
        this.hasCoverSheet = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hasCoverSheet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a>
   */
  public OlImportanceEnum getImportance() {

    String propertyName = "Importance";
    try {
      if (this.importance == null) {
        this.importance = OlImportanceEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.importance;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171439(v=office.11).aspx">InternetCodepage</a>
   */
  public Variant getInternetCodepage() {

    String propertyName = "InternetCodepage";
    try {
      if (this.internetCodepage == null) {
        this.internetCodepage = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.internetCodepage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a>
   */
  public Boolean getIsConflict() {

    String propertyName = "IsConflict";
    try {
      if (this.isConflict == null) {
        this.isConflict = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isConflict;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171446(v=office.11).aspx">IsIPFax</a>
   */
  public Variant getIsIPFax() {

    String propertyName = "IsIPFax";
    try {
      if (this.isIPFax == null) {
        this.isIPFax = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isIPFax;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a>
   */
  public ItemProperties getItemProperties() {

    String propertyName = "ItemProperties";
    try {
      if (this.itemProperties == null) {
        this.itemProperties = new ItemProperties(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.itemProperties;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a>
   */
  public Date getLastModificationTime() {

    String propertyName = "LastModificationTime";
    try {
      if (this.lastModificationTime == null) {
        this.lastModificationTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.lastModificationTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a>
   */
  public Links getLinks() {

    String propertyName = "Links";
    try {
      if (this.links == null) {
        this.links = new Links(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.links;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a>
   */
  public OlRemoteStatusEnum getMarkForDownload() {

    String propertyName = "MarkForDownload";
    try {
      if (this.markForDownload == null) {
        this.markForDownload = OlRemoteStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.markForDownload;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a>
   */
  public String getMileage() {

    String propertyName = "Mileage";
    try {
      if (this.mileage == null) {
        this.mileage = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.mileage;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a>
   */
  public Boolean getNoAging() {

    String propertyName = "NoAging";
    try {
      if (this.noAging == null) {
        this.noAging = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.noAging;
  }

  /**
   * @see <a
   *      href="http://msdn.microsoft.com/en-us/library/aa171788(v=office.11).aspx">OriginatorDeliveryReportRequested</a>
   */
  public Boolean getOriginatorDeliveryReportRequested() {

    String propertyName = "OriginatorDeliveryReportRequested";
    try {
      if (this.originatorDeliveryReportRequested == null) {
        this.originatorDeliveryReportRequested = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.originatorDeliveryReportRequested;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a>
   */
  public Variant getOutlookInternalVersion() {

    String propertyName = "OutlookInternalVersion";
    try {
      if (this.outlookInternalVersion == null) {
        this.outlookInternalVersion = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.outlookInternalVersion;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a>
   */
  public String getOutlookVersion() {

    String propertyName = "OutlookVersion";
    try {
      if (this.outlookVersion == null) {
        this.outlookVersion = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.outlookVersion;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171840(v=office.11).aspx">Permission</a>
   */
  public Variant getPermission() {

    String propertyName = "Permission";
    try {
      if (this.permission == null) {
        this.permission = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.permission;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171838(v=office.11).aspx">PermissionService</a>
   */
  public Variant getPermissionService() {

    String propertyName = "PermissionService";
    try {
      if (this.permissionService == null) {
        this.permissionService = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.permissionService;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171860(v=office.11).aspx">ReadReceiptRequested</a>
   */
  public Variant getReadReceiptRequested() {

    String propertyName = "ReadReceiptRequested";
    try {
      if (this.readReceiptRequested == null) {
        this.readReceiptRequested = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.readReceiptRequested;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171862(v=office.11).aspx">ReceivedByEntryID</a>
   */
  public String getReceivedByEntryID() {

    String propertyName = "ReceivedByEntryID";
    try {
      if (this.receivedByEntryID == null) {
        this.receivedByEntryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.receivedByEntryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171864(v=office.11).aspx">ReceivedByName</a>
   */
  public String getReceivedByName() {

    String propertyName = "ReceivedByName";
    try {
      if (this.receivedByName == null) {
        this.receivedByName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.receivedByName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171866(v=office.11).aspx">ReceivedOnBehalfOfEntryID</a>
   */
  public String getReceivedOnBehalfOfEntryID() {

    String propertyName = "ReceivedOnBehalfOfEntryID";
    try {
      if (this.receivedOnBehalfOfEntryID == null) {
        this.receivedOnBehalfOfEntryID = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.receivedOnBehalfOfEntryID;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171869(v=office.11).aspx">ReceivedOnBehalfOfName</a>
   */
  public String getReceivedOnBehalfOfName() {

    String propertyName = "ReceivedOnBehalfOfName";
    try {
      if (this.receivedOnBehalfOfName == null) {
        this.receivedOnBehalfOfName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.receivedOnBehalfOfName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171873(v=office.11).aspx">ReceivedTime</a>
   */
  public Date getReceivedTime() {

    String propertyName = "ReceivedTime";
    try {
      if (this.receivedTime == null) {
        this.receivedTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.receivedTime;
  }

  /**
   * @see <a
   *      href="http://msdn.microsoft.com/en-us/library/aa171876(v=office.11).aspx">RecipientReassignmentProhibited</a>
   */
  public Boolean getRecipientReassignmentProhibited() {

    String propertyName = "RecipientReassignmentProhibited";
    try {
      if (this.recipientReassignmentProhibited == null) {
        this.recipientReassignmentProhibited = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.recipientReassignmentProhibited;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171878(v=office.11).aspx">Recipients</a>
   */
  public Recipients getRecipients() {

    String propertyName = "Recipients";
    try {
      if (this.recipients == null) {
        this.recipients = new Recipients(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.recipients;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171894(v=office.11).aspx">ReminderOverrideDefault</a>
   */
  public Boolean getReminderOverrideDefault() {

    String propertyName = "ReminderOverrideDefault";
    try {
      if (this.reminderOverrideDefault == null) {
        this.reminderOverrideDefault = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderOverrideDefault;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171897(v=office.11).aspx">ReminderPlaySound</a>
   */
  public Boolean getReminderPlaySound() {

    String propertyName = "ReminderPlaySound";
    try {
      if (this.reminderPlaySound == null) {
        this.reminderPlaySound = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderPlaySound;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171899(v=office.11).aspx">ReminderSet</a>
   */
  public Boolean getReminderSet() {

    String propertyName = "ReminderSet";
    try {
      if (this.reminderSet == null) {
        this.reminderSet = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderSet;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171901(v=office.11).aspx">ReminderSoundFile</a>
   */
  public String getReminderSoundFile() {

    String propertyName = "ReminderSoundFile";
    try {
      if (this.reminderSoundFile == null) {
        this.reminderSoundFile = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderSoundFile;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171904(v=office.11).aspx">ReminderTime</a>
   */
  public Date getReminderTime() {

    String propertyName = "ReminderTime";
    try {
      if (this.reminderTime == null) {
        this.reminderTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171908(v=office.11).aspx">RemoteStatus</a>
   */
  public OlRemoteStatusEnum getRemoteStatus() {

    String propertyName = "RemoteStatus";
    try {
      if (this.remoteStatus == null) {
        this.remoteStatus = OlRemoteStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.remoteStatus;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171909(v=office.11).aspx">ReplyRecipientNames</a>
   */
  public String getReplyRecipientNames() {

    String propertyName = "ReplyRecipientNames";
    try {
      if (this.replyRecipientNames == null) {
        this.replyRecipientNames = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.replyRecipientNames;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171911(v=office.11).aspx">ReplyRecipients</a>
   */
  public Variant getReplyRecipients() {

    String propertyName = "ReplyRecipients";
    try {
      if (this.replyRecipients == null) {
        this.replyRecipients = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.replyRecipients;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171932(v=office.11).aspx">SaveSentMessageFolder</a>
   */
  public Variant getSaveSentMessageFolder() {

    String propertyName = "SaveSentMessageFolder";
    try {
      if (this.saveSentMessageFolder == null) {
        this.saveSentMessageFolder = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.saveSentMessageFolder;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a>
   */
  public Boolean getSaved() {

    String propertyName = "Saved";
    try {
      if (this.saved == null) {
        this.saved = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.saved;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171942(v=office.11).aspx">SenderEmailAddress</a>
   */
  public String getSenderEmailAddress() {

    String propertyName = "SenderEmailAddress";
    try {
      if (this.senderEmailAddress == null) {
        this.senderEmailAddress = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.senderEmailAddress;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171944(v=office.11).aspx">SenderEmailType</a>
   */
  public String getSenderEmailType() {

    String propertyName = "SenderEmailType";
    try {
      if (this.senderEmailType == null) {
        this.senderEmailType = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.senderEmailType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171985(v=office.11).aspx">SenderName</a>
   */
  public String getSenderName() {

    String propertyName = "SenderName";
    try {
      if (this.senderName == null) {
        this.senderName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.senderName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a>
   */
  public OlSensitivityEnum getSensitivity() {

    String propertyName = "Sensitivity";
    try {
      if (this.sensitivity == null) {
        this.sensitivity = OlSensitivityEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.sensitivity;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172010(v=office.11).aspx">Sent</a>
   */
  public Boolean getSent() {

    String propertyName = "Sent";
    try {
      if (this.sent == null) {
        this.sent = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.sent;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172005(v=office.11).aspx">SentOn</a>
   */
  public Date getSentOn() {

    String propertyName = "SentOn";
    try {
      if (this.sentOn == null) {
        this.sentOn = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.sentOn;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171998(v=office.11).aspx">SentOnBehalfOfName</a>
   */
  public String getSentOnBehalfOfName() {

    String propertyName = "SentOnBehalfOfName";
    try {
      if (this.sentOnBehalfOfName == null) {
        this.sentOnBehalfOfName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.sentOnBehalfOfName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a>
   */
  public Variant getSize() {

    String propertyName = "Size";
    try {
      if (this.size == null) {
        this.size = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.size;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a>
   */
  public String getSubject() {

    String propertyName = "Subject";
    try {
      if (this.subject == null) {
        this.subject = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.subject;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220151(v=office.11).aspx">Submitted</a>
   */
  public Boolean getSubmitted() {

    String propertyName = "Submitted";
    try {
      if (this.submitted == null) {
        this.submitted = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.submitted;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220513(v=office.11).aspx">To</a>
   */
  public String getTo() {

    String propertyName = "To";
    try {
      if (this.to == null) {
        this.to = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.to;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a>
   */
  public Boolean getUnRead() {

    String propertyName = "UnRead";
    try {
      if (this.unRead == null) {
        this.unRead = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.unRead;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a>
   */
  public UserProperties getUserProperties() {

    String propertyName = "UserProperties";
    try {
      if (this.userProperties == null) {
        this.userProperties = new UserProperties(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.userProperties;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221825(v=office.11).aspx">VotingOptions</a>
   */
  public String getVotingOptions() {

    String propertyName = "VotingOptions";
    try {
      if (this.votingOptions == null) {
        this.votingOptions = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.votingOptions;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221831(v=office.11).aspx">VotingResponse</a>
   */
  public String getVotingResponse() {

    String propertyName = "VotingResponse";
    try {
      if (this.votingResponse == null) {
        this.votingResponse = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.votingResponse;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.actions != null) {
      this.actions.dispose();
    }
    if (this.attachments != null) {
      this.attachments.dispose();
    }
    if (this.conflicts != null) {
      this.conflicts.dispose();
    }
    if (this.formDescription != null) {
      this.formDescription.dispose();
    }
    if (this.itemProperties != null) {
      this.itemProperties.dispose();
    }
    if (this.links != null) {
      this.links.dispose();
    }
    if (this.recipients != null) {
      this.recipients.dispose();
    }
    if (this.userProperties != null) {
      this.userProperties.dispose();
    }
  }

}
