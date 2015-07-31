package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211067(v=office.11).aspx">TaskItem</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211349(v=office.11).aspx">ActualWork</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211786(v=office.11).aspx">CardData</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211804(v=office.11).aspx">Complete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211810(v=office.11).aspx">ContactNames</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211825(v=office.11).aspx">DateCompleted</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211831(v=office.11).aspx">DelegationState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211832(v=office.11).aspx">Delegator</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211846(v=office.11).aspx">DueDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171439(v=office.11).aspx">InternetCodepage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171449(v=office.11).aspx">IsRecurring</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171783(v=office.11).aspx">Ordinal</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171809(v=office.11).aspx">Owner</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171806(v=office.11).aspx">Ownership</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171834(v=office.11).aspx">PercentComplete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171878(v=office.11).aspx">Recipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171894(v=office.11).aspx">ReminderOverrideDefault</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171897(v=office.11).aspx">ReminderPlaySound</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171899(v=office.11).aspx">ReminderSet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171901(v=office.11).aspx">ReminderSoundFile</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171904(v=office.11).aspx">ReminderTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171921(v=office.11).aspx">ResponseState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171927(v=office.11).aspx">Role</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171933(v=office.11).aspx">SchedulePlusPriority</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172053(v=office.11).aspx">StartDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212419(v=office.11).aspx">Status</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212413(v=office.11).aspx">StatusOnCompletionRecipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212416(v=office.11).aspx">StatusUpdateRecipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220284(v=office.11).aspx">TeamTask</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220432(v=office.11).aspx">TotalWork</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220073(v=office.11).aspx">Assign</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220074(v=office.11).aspx">CancelResponseState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220076(v=office.11).aspx">ClearRecurrencePattern</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220080(v=office.11).aspx">Copy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220115(v=office.11).aspx">GetRecurrencePattern</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220125(v=office.11).aspx">MarkComplete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220127(v=office.11).aspx">Move</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220131(v=office.11).aspx">PrintOut</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210273(v=office.11).aspx">Respond</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210281(v=office.11).aspx">Save</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210279(v=office.11).aspx">SaveAs</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210285(v=office.11).aspx">Send</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210294(v=office.11).aspx">ShowCategoriesDialog</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210300(v=office.11).aspx">SkipRecurrence</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210310(v=office.11).aspx">StatusReport</a>
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
 *      href="http://msdn.microsoft.com/en-us/library/aa210995(v=office.11).aspx">Recipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211095(v=office.11).aspx">UserProperties</a>
 *      </p>*
 * @author eugeis
 */

public class TaskItem extends OleAuto {

  protected Actions actions;

  protected Variant actualWork;

  protected Attachments attachments;

  protected Boolean autoResolvedWinner;

  protected String billingInformation;

  protected String body;

  protected String cardData;

  protected String categories;

  protected String companies;

  protected Boolean complete;

  protected Conflicts conflicts;

  protected String contactNames;

  protected String conversationIndex;

  protected String conversationTopic;

  protected Date creationTime;

  protected Date dateCompleted;

  protected OlTaskDelegationStateEnum delegationState;

  protected String delegator;

  protected OlDownloadStateEnum downloadState;

  protected Date dueDate;

  protected String entryID;

  protected FormDescription formDescription;

  protected Variant getInspector;

  protected OlImportanceEnum importance;

  protected Variant internetCodepage;

  protected Boolean isConflict;

  protected Boolean isRecurring;

  protected ItemProperties itemProperties;

  protected Date lastModificationTime;

  protected Links links;

  protected OlRemoteStatusEnum markForDownload;

  protected String messageClass;

  protected String mileage;

  protected Boolean noAging;

  protected Variant ordinal;

  protected Variant outlookInternalVersion;

  protected String outlookVersion;

  protected String owner;

  protected OlTaskOwnershipEnum ownership;

  protected Variant percentComplete;

  protected Recipients recipients;

  protected Boolean reminderOverrideDefault;

  protected Boolean reminderPlaySound;

  protected Boolean reminderSet;

  protected String reminderSoundFile;

  protected Date reminderTime;

  protected OlTaskResponseEnum responseState;

  protected String role;

  protected Boolean saved;

  protected String schedulePlusPriority;

  protected OlSensitivityEnum sensitivity;

  protected Variant size;

  protected Date startDate;

  protected OlTaskStatusEnum status;

  protected String statusOnCompletionRecipients;

  protected String statusUpdateRecipients;

  protected String subject;

  protected Boolean teamTask;

  protected Variant totalWork;

  protected Boolean unRead;

  protected UserProperties userProperties;

  public TaskItem(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getActions();
    getActualWork();
    getAttachments();
    getAutoResolvedWinner();
    getBillingInformation();
    getBody();
    getCardData();
    getCategories();
    getCompanies();
    getComplete();
    getConflicts();
    getContactNames();
    getConversationIndex();
    getConversationTopic();
    getCreationTime();
    getDateCompleted();
    getDelegationState();
    getDelegator();
    getDownloadState();
    getDueDate();
    getEntryID();
    getFormDescription();
    getGetInspector();
    getImportance();
    getInternetCodepage();
    getIsConflict();
    getIsRecurring();
    getItemProperties();
    getLastModificationTime();
    getLinks();
    getMarkForDownload();
    getMessageClass();
    getMileage();
    getNoAging();
    getOrdinal();
    getOutlookInternalVersion();
    getOutlookVersion();
    getOwner();
    getOwnership();
    getPercentComplete();
    getRecipients();
    getReminderOverrideDefault();
    getReminderPlaySound();
    getReminderSet();
    getReminderSoundFile();
    getReminderTime();
    getResponseState();
    getRole();
    getSaved();
    getSchedulePlusPriority();
    getSensitivity();
    getSize();
    getStartDate();
    getStatus();
    getStatusOnCompletionRecipients();
    getStatusUpdateRecipients();
    getSubject();
    getTeamTask();
    getTotalWork();
    getUnRead();
    getUserProperties();
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211349(v=office.11).aspx">ActualWork</a>
   */
  public Variant getActualWork() {

    String propertyName = "ActualWork";
    try {
      if (this.actualWork == null) {
        this.actualWork = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.actualWork;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211786(v=office.11).aspx">CardData</a>
   */
  public String getCardData() {

    String propertyName = "CardData";
    try {
      if (this.cardData == null) {
        this.cardData = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.cardData;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211804(v=office.11).aspx">Complete</a>
   */
  public Boolean getComplete() {

    String propertyName = "Complete";
    try {
      if (this.complete == null) {
        this.complete = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.complete;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211810(v=office.11).aspx">ContactNames</a>
   */
  public String getContactNames() {

    String propertyName = "ContactNames";
    try {
      if (this.contactNames == null) {
        this.contactNames = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.contactNames;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211825(v=office.11).aspx">DateCompleted</a>
   */
  public Date getDateCompleted() {

    String propertyName = "DateCompleted";
    try {
      if (this.dateCompleted == null) {
        this.dateCompleted = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.dateCompleted;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211831(v=office.11).aspx">DelegationState</a>
   */
  public OlTaskDelegationStateEnum getDelegationState() {

    String propertyName = "DelegationState";
    try {
      if (this.delegationState == null) {
        this.delegationState = OlTaskDelegationStateEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.delegationState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211832(v=office.11).aspx">Delegator</a>
   */
  public String getDelegator() {

    String propertyName = "Delegator";
    try {
      if (this.delegator == null) {
        this.delegator = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.delegator;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211846(v=office.11).aspx">DueDate</a>
   */
  public Date getDueDate() {

    String propertyName = "DueDate";
    try {
      if (this.dueDate == null) {
        this.dueDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.dueDate;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171449(v=office.11).aspx">IsRecurring</a>
   */
  public Boolean getIsRecurring() {

    String propertyName = "IsRecurring";
    try {
      if (this.isRecurring == null) {
        this.isRecurring = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isRecurring;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171783(v=office.11).aspx">Ordinal</a>
   */
  public Variant getOrdinal() {

    String propertyName = "Ordinal";
    try {
      if (this.ordinal == null) {
        this.ordinal = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.ordinal;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171809(v=office.11).aspx">Owner</a>
   */
  public String getOwner() {

    String propertyName = "Owner";
    try {
      if (this.owner == null) {
        this.owner = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.owner;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171806(v=office.11).aspx">Ownership</a>
   */
  public OlTaskOwnershipEnum getOwnership() {

    String propertyName = "Ownership";
    try {
      if (this.ownership == null) {
        this.ownership = OlTaskOwnershipEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.ownership;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171834(v=office.11).aspx">PercentComplete</a>
   */
  public Variant getPercentComplete() {

    String propertyName = "PercentComplete";
    try {
      if (this.percentComplete == null) {
        this.percentComplete = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.percentComplete;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171921(v=office.11).aspx">ResponseState</a>
   */
  public OlTaskResponseEnum getResponseState() {

    String propertyName = "ResponseState";
    try {
      if (this.responseState == null) {
        this.responseState = OlTaskResponseEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.responseState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171927(v=office.11).aspx">Role</a>
   */
  public String getRole() {

    String propertyName = "Role";
    try {
      if (this.role == null) {
        this.role = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.role;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171933(v=office.11).aspx">SchedulePlusPriority</a>
   */
  public String getSchedulePlusPriority() {

    String propertyName = "SchedulePlusPriority";
    try {
      if (this.schedulePlusPriority == null) {
        this.schedulePlusPriority = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.schedulePlusPriority;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172053(v=office.11).aspx">StartDate</a>
   */
  public Date getStartDate() {

    String propertyName = "StartDate";
    try {
      if (this.startDate == null) {
        this.startDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.startDate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212419(v=office.11).aspx">Status</a>
   */
  public OlTaskStatusEnum getStatus() {

    String propertyName = "Status";
    try {
      if (this.status == null) {
        this.status = OlTaskStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.status;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212413(v=office.11).aspx">StatusOnCompletionRecipients</a>
   */
  public String getStatusOnCompletionRecipients() {

    String propertyName = "StatusOnCompletionRecipients";
    try {
      if (this.statusOnCompletionRecipients == null) {
        this.statusOnCompletionRecipients = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.statusOnCompletionRecipients;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212416(v=office.11).aspx">StatusUpdateRecipients</a>
   */
  public String getStatusUpdateRecipients() {

    String propertyName = "StatusUpdateRecipients";
    try {
      if (this.statusUpdateRecipients == null) {
        this.statusUpdateRecipients = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.statusUpdateRecipients;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220284(v=office.11).aspx">TeamTask</a>
   */
  public Boolean getTeamTask() {

    String propertyName = "TeamTask";
    try {
      if (this.teamTask == null) {
        this.teamTask = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.teamTask;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220432(v=office.11).aspx">TotalWork</a>
   */
  public Variant getTotalWork() {

    String propertyName = "TotalWork";
    try {
      if (this.totalWork == null) {
        this.totalWork = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.totalWork;
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
