package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210899(v=office.11).aspx">AppointmentItem</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211346(v=office.11).aspx">Actions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211368(v=office.11).aspx">AllDayEvent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211408(v=office.11).aspx">Attachments</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211414(v=office.11).aspx">AutoResolvedWinner</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211425(v=office.11).aspx">BillingInformation</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211433(v=office.11).aspx">Body</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211783(v=office.11).aspx">BusyStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211788(v=office.11).aspx">Categories</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211798(v=office.11).aspx">Companies</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211806(v=office.11).aspx">ConferenceServerAllowExternal</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211807(v=office.11).aspx">ConferenceServerPassword</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211808(v=office.11).aspx">Conflicts</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211813(v=office.11).aspx">ConversationIndex</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211814(v=office.11).aspx">ConversationTopic</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211817(v=office.11).aspx">CreationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211845(v=office.11).aspx">DownloadState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211847(v=office.11).aspx">Duration</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211864(v=office.11).aspx">End</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211865(v=office.11).aspx">EntryID</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212024(v=office.11).aspx">FormDescription</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212040(v=office.11).aspx">GetInspector</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171432(v=office.11).aspx">Importance</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171439(v=office.11).aspx">InternetCodepage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171442(v=office.11).aspx">IsConflict</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171447(v=office.11).aspx">IsOnlineMeeting</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171449(v=office.11).aspx">IsRecurring</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171454(v=office.11).aspx">ItemProperties</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171467(v=office.11).aspx">LastModificationTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171471(v=office.11).aspx">Links</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171472(v=office.11).aspx">Location</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171484(v=office.11).aspx">MarkForDownload</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171486(v=office.11).aspx">MeetingStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171487(v=office.11).aspx">MeetingWorkspaceURL</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171490(v=office.11).aspx">MessageClass</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171492(v=office.11).aspx">Mileage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171697(v=office.11).aspx">NetMeetingAutoStart</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171701(v=office.11).aspx">NetMeetingDocPathName</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171702(v=office.11).aspx">NetMeetingOrganizerAlias</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171706(v=office.11).aspx">NetMeetingServer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171712(v=office.11).aspx">NetMeetingType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171716(v=office.11).aspx">NetShowURL</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171769(v=office.11).aspx">NoAging</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171782(v=office.11).aspx">OptionalAttendees</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171785(v=office.11).aspx">Organizer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171803(v=office.11).aspx">OutlookInternalVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171805(v=office.11).aspx">OutlookVersion</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171878(v=office.11).aspx">Recipients</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171880(v=office.11).aspx">RecurrenceState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171892(v=office.11).aspx">ReminderMinutesBeforeStart</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171894(v=office.11).aspx">ReminderOverrideDefault</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171897(v=office.11).aspx">ReminderPlaySound</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171899(v=office.11).aspx">ReminderSet</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171901(v=office.11).aspx">ReminderSoundFile</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171914(v=office.11).aspx">ReplyTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171915(v=office.11).aspx">RequiredAttendees</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171918(v=office.11).aspx">Resources</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171919(v=office.11).aspx">ResponseRequested</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171922(v=office.11).aspx">ResponseStatus</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171928(v=office.11).aspx">Saved</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171987(v=office.11).aspx">Sensitivity</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172040(v=office.11).aspx">Size</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212409(v=office.11).aspx">Start</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212425(v=office.11).aspx">Subject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221742(v=office.11).aspx">UnRead</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221771(v=office.11).aspx">UserProperties</a>
 *      </p>
 *      <p>
 *      Methods | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220076(v=office.11).aspx">ClearRecurrencePattern</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220080(v=office.11).aspx">Copy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220085(v=office.11).aspx">Delete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220094(v=office.11).aspx">ForwardAsVcal</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220115(v=office.11).aspx">GetRecurrencePattern</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220127(v=office.11).aspx">Move</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220131(v=office.11).aspx">PrintOut</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210273(v=office.11).aspx">Respond</a> | <a
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
 *      Parent Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210913(v=office.11).aspx">Exception</a>
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

public class AppointmentItem extends OleAuto {

  protected Actions actions;

  protected Boolean allDayEvent;

  protected Attachments attachments;

  protected Boolean autoResolvedWinner;

  protected String billingInformation;

  protected String body;

  protected OlBusyStatusEnum busyStatus;

  protected String categories;

  protected String companies;

  protected Variant conferenceServerAllowExternal;

  protected Variant conferenceServerPassword;

  protected Conflicts conflicts;

  protected String conversationIndex;

  protected String conversationTopic;

  protected Date creationTime;

  protected OlDownloadStateEnum downloadState;

  protected Variant duration;

  protected Date end;

  protected String entryID;

  protected FormDescription formDescription;

  protected Variant getInspector;

  protected OlImportanceEnum importance;

  protected Variant internetCodepage;

  protected Boolean isConflict;

  protected Boolean isOnlineMeeting;

  protected Boolean isRecurring;

  protected ItemProperties itemProperties;

  protected Date lastModificationTime;

  protected Links links;

  protected String location;

  protected OlRemoteStatusEnum markForDownload;

  protected OlMeetingStatusEnum meetingStatus;

  protected Variant meetingWorkspaceURL;

  protected String messageClass;

  protected String mileage;

  protected Boolean netMeetingAutoStart;

  protected String netMeetingDocPathName;

  protected String netMeetingOrganizerAlias;

  protected String netMeetingServer;

  protected OlNetMeetingTypeEnum netMeetingType;

  protected String netShowURL;

  protected Boolean noAging;

  protected String optionalAttendees;

  protected String organizer;

  protected Variant outlookInternalVersion;

  protected String outlookVersion;

  protected Recipients recipients;

  protected OlRecurrenceStateEnum recurrenceState;

  protected Variant reminderMinutesBeforeStart;

  protected Boolean reminderOverrideDefault;

  protected Boolean reminderPlaySound;

  protected Boolean reminderSet;

  protected String reminderSoundFile;

  protected Date replyTime;

  protected String requiredAttendees;

  protected String resources;

  protected Boolean responseRequested;

  protected OlResponseStatusEnum responseStatus;

  protected Boolean saved;

  protected OlSensitivityEnum sensitivity;

  protected Variant size;

  protected Date start;

  protected String subject;

  protected Boolean unRead;

  protected UserProperties userProperties;

  public AppointmentItem(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getActions();
    getAllDayEvent();
    getAttachments();
    getAutoResolvedWinner();
    getBillingInformation();
    getBody();
    getBusyStatus();
    getCategories();
    getCompanies();
    getConferenceServerAllowExternal();
    getConferenceServerPassword();
    getConflicts();
    getConversationIndex();
    getConversationTopic();
    getCreationTime();
    getDownloadState();
    getDuration();
    getEnd();
    getEntryID();
    getFormDescription();
    getGetInspector();
    getImportance();
    getInternetCodepage();
    getIsConflict();
    getIsOnlineMeeting();
    getIsRecurring();
    getItemProperties();
    getLastModificationTime();
    getLinks();
    getLocation();
    getMarkForDownload();
    getMeetingStatus();
    getMeetingWorkspaceURL();
    getMessageClass();
    getMileage();
    getNetMeetingAutoStart();
    getNetMeetingDocPathName();
    getNetMeetingOrganizerAlias();
    getNetMeetingServer();
    getNetMeetingType();
    getNetShowURL();
    getNoAging();
    getOptionalAttendees();
    getOrganizer();
    getOutlookInternalVersion();
    getOutlookVersion();
    getRecipients();
    getRecurrenceState();
    getReminderMinutesBeforeStart();
    getReminderOverrideDefault();
    getReminderPlaySound();
    getReminderSet();
    getReminderSoundFile();
    getReplyTime();
    getRequiredAttendees();
    getResources();
    getResponseRequested();
    getResponseStatus();
    getSaved();
    getSensitivity();
    getSize();
    getStart();
    getSubject();
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211368(v=office.11).aspx">AllDayEvent</a>
   */
  public Boolean getAllDayEvent() {

    String propertyName = "AllDayEvent";
    try {
      if (this.allDayEvent == null) {
        this.allDayEvent = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.allDayEvent;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211783(v=office.11).aspx">BusyStatus</a>
   */
  public OlBusyStatusEnum getBusyStatus() {

    String propertyName = "BusyStatus";
    try {
      if (this.busyStatus == null) {
        this.busyStatus = OlBusyStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.busyStatus;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211806(v=office.11).aspx">ConferenceServerAllowExternal</a>
   */
  public Variant getConferenceServerAllowExternal() {

    String propertyName = "ConferenceServerAllowExternal";
    try {
      if (this.conferenceServerAllowExternal == null) {
        this.conferenceServerAllowExternal = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conferenceServerAllowExternal;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211807(v=office.11).aspx">ConferenceServerPassword</a>
   */
  public Variant getConferenceServerPassword() {

    String propertyName = "ConferenceServerPassword";
    try {
      if (this.conferenceServerPassword == null) {
        this.conferenceServerPassword = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.conferenceServerPassword;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211847(v=office.11).aspx">Duration</a>
   */
  public Variant getDuration() {

    String propertyName = "Duration";
    try {
      if (this.duration == null) {
        this.duration = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.duration;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211864(v=office.11).aspx">End</a>
   */
  public Date getEnd() {

    String propertyName = "End";
    try {
      if (this.end == null) {
        this.end = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.end;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171447(v=office.11).aspx">IsOnlineMeeting</a>
   */
  public Boolean getIsOnlineMeeting() {

    String propertyName = "IsOnlineMeeting";
    try {
      if (this.isOnlineMeeting == null) {
        this.isOnlineMeeting = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isOnlineMeeting;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171472(v=office.11).aspx">Location</a>
   */
  public String getLocation() {

    String propertyName = "Location";
    try {
      if (this.location == null) {
        this.location = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.location;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171486(v=office.11).aspx">MeetingStatus</a>
   */
  public OlMeetingStatusEnum getMeetingStatus() {

    String propertyName = "MeetingStatus";
    try {
      if (this.meetingStatus == null) {
        this.meetingStatus = OlMeetingStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.meetingStatus;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171487(v=office.11).aspx">MeetingWorkspaceURL</a>
   */
  public Variant getMeetingWorkspaceURL() {

    String propertyName = "MeetingWorkspaceURL";
    try {
      if (this.meetingWorkspaceURL == null) {
        this.meetingWorkspaceURL = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.meetingWorkspaceURL;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171697(v=office.11).aspx">NetMeetingAutoStart</a>
   */
  public Boolean getNetMeetingAutoStart() {

    String propertyName = "NetMeetingAutoStart";
    try {
      if (this.netMeetingAutoStart == null) {
        this.netMeetingAutoStart = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingAutoStart;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171701(v=office.11).aspx">NetMeetingDocPathName</a>
   */
  public String getNetMeetingDocPathName() {

    String propertyName = "NetMeetingDocPathName";
    try {
      if (this.netMeetingDocPathName == null) {
        this.netMeetingDocPathName = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingDocPathName;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171702(v=office.11).aspx">NetMeetingOrganizerAlias</a>
   */
  public String getNetMeetingOrganizerAlias() {

    String propertyName = "NetMeetingOrganizerAlias";
    try {
      if (this.netMeetingOrganizerAlias == null) {
        this.netMeetingOrganizerAlias = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingOrganizerAlias;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171706(v=office.11).aspx">NetMeetingServer</a>
   */
  public String getNetMeetingServer() {

    String propertyName = "NetMeetingServer";
    try {
      if (this.netMeetingServer == null) {
        this.netMeetingServer = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingServer;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171712(v=office.11).aspx">NetMeetingType</a>
   */
  public OlNetMeetingTypeEnum getNetMeetingType() {

    String propertyName = "NetMeetingType";
    try {
      if (this.netMeetingType == null) {
        this.netMeetingType = OlNetMeetingTypeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netMeetingType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171716(v=office.11).aspx">NetShowURL</a>
   */
  public String getNetShowURL() {

    String propertyName = "NetShowURL";
    try {
      if (this.netShowURL == null) {
        this.netShowURL = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.netShowURL;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171782(v=office.11).aspx">OptionalAttendees</a>
   */
  public String getOptionalAttendees() {

    String propertyName = "OptionalAttendees";
    try {
      if (this.optionalAttendees == null) {
        this.optionalAttendees = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.optionalAttendees;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171785(v=office.11).aspx">Organizer</a>
   */
  public String getOrganizer() {

    String propertyName = "Organizer";
    try {
      if (this.organizer == null) {
        this.organizer = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.organizer;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171880(v=office.11).aspx">RecurrenceState</a>
   */
  public OlRecurrenceStateEnum getRecurrenceState() {

    String propertyName = "RecurrenceState";
    try {
      if (this.recurrenceState == null) {
        this.recurrenceState = OlRecurrenceStateEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.recurrenceState;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171892(v=office.11).aspx">ReminderMinutesBeforeStart</a>
   */
  public Variant getReminderMinutesBeforeStart() {

    String propertyName = "ReminderMinutesBeforeStart";
    try {
      if (this.reminderMinutesBeforeStart == null) {
        this.reminderMinutesBeforeStart = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminderMinutesBeforeStart;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171914(v=office.11).aspx">ReplyTime</a>
   */
  public Date getReplyTime() {

    String propertyName = "ReplyTime";
    try {
      if (this.replyTime == null) {
        this.replyTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.replyTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171915(v=office.11).aspx">RequiredAttendees</a>
   */
  public String getRequiredAttendees() {

    String propertyName = "RequiredAttendees";
    try {
      if (this.requiredAttendees == null) {
        this.requiredAttendees = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.requiredAttendees;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171918(v=office.11).aspx">Resources</a>
   */
  public String getResources() {

    String propertyName = "Resources";
    try {
      if (this.resources == null) {
        this.resources = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.resources;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171919(v=office.11).aspx">ResponseRequested</a>
   */
  public Boolean getResponseRequested() {

    String propertyName = "ResponseRequested";
    try {
      if (this.responseRequested == null) {
        this.responseRequested = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.responseRequested;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171922(v=office.11).aspx">ResponseStatus</a>
   */
  public OlResponseStatusEnum getResponseStatus() {

    String propertyName = "ResponseStatus";
    try {
      if (this.responseStatus == null) {
        this.responseStatus = OlResponseStatusEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.responseStatus;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212409(v=office.11).aspx">Start</a>
   */
  public Date getStart() {

    String propertyName = "Start";
    try {
      if (this.start == null) {
        this.start = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.start;
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
