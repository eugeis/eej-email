package ee.email.outlook.base;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa219371(v=office.11).aspx">OlDefaultFolders</a>
 *      </p>
 * @author eugeis
 */
public enum OlDefaultFoldersEnum {
  olFolderDeletedItems(3), olFolderOutbox(4), olFolderSentMail(5), olFolderInbox(6), olFolderCalendar(9), olFolderContacts(10), olFolderJournal(11), olFolderNotes(12), olFolderTasks(13), olFolderDrafts(16), olFolderConflicts(19), olPublicFoldersAllPublicFolders(18), olFolderLocalFailures(21), olFolderSyncIssues(20), olFolderJunk(23), olFolderServerFailures(22);

  private final int value;

  private OlDefaultFoldersEnum(int value) {

    this.value = value;
  }

  public int getValue() {

    return this.value;
  }

  public static OlDefaultFoldersEnum findEnum(Integer value) {

    if (value != null) {
      for (OlDefaultFoldersEnum objEnum : values()) {
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

  public boolean isOlFolderDeletedItems() {

    return olFolderDeletedItems == this;
  }

  public boolean isOlFolderOutbox() {

    return olFolderOutbox == this;
  }

  public boolean isOlFolderSentMail() {

    return olFolderSentMail == this;
  }

  public boolean isOlFolderInbox() {

    return olFolderInbox == this;
  }

  public boolean isOlFolderCalendar() {

    return olFolderCalendar == this;
  }

  public boolean isOlFolderContacts() {

    return olFolderContacts == this;
  }

  public boolean isOlFolderJournal() {

    return olFolderJournal == this;
  }

  public boolean isOlFolderNotes() {

    return olFolderNotes == this;
  }

  public boolean isOlFolderTasks() {

    return olFolderTasks == this;
  }

  public boolean isOlFolderDrafts() {

    return olFolderDrafts == this;
  }

  public boolean isOlFolderConflicts() {

    return olFolderConflicts == this;
  }

  public boolean isOlPublicFoldersAllPublicFolders() {

    return olPublicFoldersAllPublicFolders == this;
  }

  public boolean isOlFolderLocalFailures() {

    return olFolderLocalFailures == this;
  }

  public boolean isOlFolderSyncIssues() {

    return olFolderSyncIssues == this;
  }

  public boolean isOlFolderJunk() {

    return olFolderJunk == this;
  }

  public boolean isOlFolderServerFailures() {

    return olFolderServerFailures == this;
  }
}
