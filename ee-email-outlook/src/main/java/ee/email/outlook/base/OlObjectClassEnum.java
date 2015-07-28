package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlObjectClass</a>
 *      </p>
 * @author eugeis
 */
public enum OlObjectClassEnum {
  olApplication(0), olNamespace(1), olFolder(2), olRecipient(4), olAttachment(5), olAddressList(7), olAddressEntry(8), olFolders(
      15), olRecipients(17), olItems(16), olAttachments(18), olAddressEntries(21), olAddressLists(20), olAppointment(26), olExceptions(
      29), olRecurrencePattern(28), olException(30), olExplorer(34), olInspector(35), olAction(32), olActions(33), olUserProperties(
      38), olUserProperty(39), olPages(36), olFormDescription(37), olJournal(42), olMail(43), olContact(40), olDocument(
      41), olReport(46), olRemote(47), olNote(44), olPost(45), olTaskRequestAccept(51), olTaskRequestUpdate(50), olTaskRequest(
      49), olTask(48), olMeetingResponseNegative(55), olMeetingCancellation(54), olMeetingRequest(53), olTaskRequestDecline(
      52), olMeetingResponseTentative(57), olMeetingResponsePositive(56), olOutlookBarPane(63), olPanes(62), olInspectors(
      61), olExplorers(60), olOutlookBarShortcut(68), olDistributionList(69), olPropertyPageSite(70), olPropertyPages(
      71), olOutlookBarStorage(64), olOutlookBarGroups(65), olOutlookBarGroup(66), olOutlookBarShortcuts(67), olLinks(
      76), olSearch(77), olResults(78), olViews(79), olSyncObject(72), olSyncObjects(73), olSelection(74), olLink(75), olView(
      80), olReminders(100), olReminder(101), olItemProperties(98), olItemProperty(99), olConflicts(118), olConflict(
      117);

  private final int value;

  private OlObjectClassEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlObjectClassEnum findEnum(Integer value) {

    if (value != null) {
      for (OlObjectClassEnum objEnum : values()) {
        if (objEnum.value == value) {
          return objEnum;
        }
      }
    }
    return null;
  }

  public boolean isValue(int value) {

    return this.value == value;
  }

  public boolean isOlApplication() {

    return olApplication == this;
  }

  public boolean isOlNamespace() {

    return olNamespace == this;
  }

  public boolean isOlFolder() {

    return olFolder == this;
  }

  public boolean isOlRecipient() {

    return olRecipient == this;
  }

  public boolean isOlAttachment() {

    return olAttachment == this;
  }

  public boolean isOlAddressList() {

    return olAddressList == this;
  }

  public boolean isOlAddressEntry() {

    return olAddressEntry == this;
  }

  public boolean isOlFolders() {

    return olFolders == this;
  }

  public boolean isOlRecipients() {

    return olRecipients == this;
  }

  public boolean isOlItems() {

    return olItems == this;
  }

  public boolean isOlAttachments() {

    return olAttachments == this;
  }

  public boolean isOlAddressEntries() {

    return olAddressEntries == this;
  }

  public boolean isOlAddressLists() {

    return olAddressLists == this;
  }

  public boolean isOlAppointment() {

    return olAppointment == this;
  }

  public boolean isOlExceptions() {

    return olExceptions == this;
  }

  public boolean isOlRecurrencePattern() {

    return olRecurrencePattern == this;
  }

  public boolean isOlException() {

    return olException == this;
  }

  public boolean isOlExplorer() {

    return olExplorer == this;
  }

  public boolean isOlInspector() {

    return olInspector == this;
  }

  public boolean isOlAction() {

    return olAction == this;
  }

  public boolean isOlActions() {

    return olActions == this;
  }

  public boolean isOlUserProperties() {

    return olUserProperties == this;
  }

  public boolean isOlUserProperty() {

    return olUserProperty == this;
  }

  public boolean isOlPages() {

    return olPages == this;
  }

  public boolean isOlFormDescription() {

    return olFormDescription == this;
  }

  public boolean isOlJournal() {

    return olJournal == this;
  }

  public boolean isOlMail() {

    return olMail == this;
  }

  public boolean isOlContact() {

    return olContact == this;
  }

  public boolean isOlDocument() {

    return olDocument == this;
  }

  public boolean isOlReport() {

    return olReport == this;
  }

  public boolean isOlRemote() {

    return olRemote == this;
  }

  public boolean isOlNote() {

    return olNote == this;
  }

  public boolean isOlPost() {

    return olPost == this;
  }

  public boolean isOlTaskRequestAccept() {

    return olTaskRequestAccept == this;
  }

  public boolean isOlTaskRequestUpdate() {

    return olTaskRequestUpdate == this;
  }

  public boolean isOlTaskRequest() {

    return olTaskRequest == this;
  }

  public boolean isOlTask() {

    return olTask == this;
  }

  public boolean isOlMeetingResponseNegative() {

    return olMeetingResponseNegative == this;
  }

  public boolean isOlMeetingCancellation() {

    return olMeetingCancellation == this;
  }

  public boolean isOlMeetingRequest() {

    return olMeetingRequest == this;
  }

  public boolean isOlTaskRequestDecline() {

    return olTaskRequestDecline == this;
  }

  public boolean isOlMeetingResponseTentative() {

    return olMeetingResponseTentative == this;
  }

  public boolean isOlMeetingResponsePositive() {

    return olMeetingResponsePositive == this;
  }

  public boolean isOlOutlookBarPane() {

    return olOutlookBarPane == this;
  }

  public boolean isOlPanes() {

    return olPanes == this;
  }

  public boolean isOlInspectors() {

    return olInspectors == this;
  }

  public boolean isOlExplorers() {

    return olExplorers == this;
  }

  public boolean isOlOutlookBarShortcut() {

    return olOutlookBarShortcut == this;
  }

  public boolean isOlDistributionList() {

    return olDistributionList == this;
  }

  public boolean isOlPropertyPageSite() {

    return olPropertyPageSite == this;
  }

  public boolean isOlPropertyPages() {

    return olPropertyPages == this;
  }

  public boolean isOlOutlookBarStorage() {

    return olOutlookBarStorage == this;
  }

  public boolean isOlOutlookBarGroups() {

    return olOutlookBarGroups == this;
  }

  public boolean isOlOutlookBarGroup() {

    return olOutlookBarGroup == this;
  }

  public boolean isOlOutlookBarShortcuts() {

    return olOutlookBarShortcuts == this;
  }

  public boolean isOlLinks() {

    return olLinks == this;
  }

  public boolean isOlSearch() {

    return olSearch == this;
  }

  public boolean isOlResults() {

    return olResults == this;
  }

  public boolean isOlViews() {

    return olViews == this;
  }

  public boolean isOlSyncObject() {

    return olSyncObject == this;
  }

  public boolean isOlSyncObjects() {

    return olSyncObjects == this;
  }

  public boolean isOlSelection() {

    return olSelection == this;
  }

  public boolean isOlLink() {

    return olLink == this;
  }

  public boolean isOlView() {

    return olView == this;
  }

  public boolean isOlReminders() {

    return olReminders == this;
  }

  public boolean isOlReminder() {

    return olReminder == this;
  }

  public boolean isOlItemProperties() {

    return olItemProperties == this;
  }

  public boolean isOlItemProperty() {

    return olItemProperty == this;
  }

  public boolean isOlConflicts() {

    return olConflicts == this;
  }

  public boolean isOlConflict() {

    return olConflict == this;
  }
}
