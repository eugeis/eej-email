package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.base.AppointmentItem;
import ee.email.outlook.base.ContactItem;
import ee.email.outlook.base.DocumentItem;
import ee.email.outlook.base.JournalItem;
import ee.email.outlook.base.MailItem;
import ee.email.outlook.base.MeetingItem;
import ee.email.outlook.base.NoteItem;
import ee.email.outlook.base.OlObjectClassEnum;
import ee.email.outlook.base.PostItem;
import ee.email.outlook.base.RemoteItem;
import ee.email.outlook.base.ReportItem;
import ee.email.outlook.base.TaskItem;
import ee.email.outlook.base.TaskRequestAcceptItem;
import ee.email.outlook.base.TaskRequestDeclineItem;
import ee.email.outlook.base.TaskRequestUpdateItem;

public class ItemOleAutoFactory<E extends OleAuto> implements OleAutoFactory<E> {

  @SuppressWarnings("unchecked")
  @Override
  public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

    E ret = (E) new OleAuto(auto, initImmediate);
    OlObjectClassEnum oleClass = ret.getClassOle();
    if (oleClass != null) {
      if (oleClass.isOlMail()) {
        ret = (E) new MailItem(auto, initImmediate);
      } else if (oleClass.isOlAppointment()) {
        ret = (E) new AppointmentItem(auto, initImmediate);
      } else if (oleClass.isOlContact()) {
        ret = (E) new ContactItem(auto, initImmediate);
      } else if (oleClass.isOlDocument()) {
        ret = (E) new DocumentItem(auto, initImmediate);
      } else if (oleClass.isOlJournal()) {
        ret = (E) new JournalItem(auto, initImmediate);
      } else if (oleClass.isOlMeetingRequest()) {
        ret = (E) new MeetingItem(auto, initImmediate);
      } else if (oleClass.isOlNote()) {
        ret = (E) new NoteItem(auto, initImmediate);
      } else if (oleClass.isOlPost()) {
        ret = (E) new PostItem(auto, initImmediate);
      } else if (oleClass.isOlRemote()) {
        ret = (E) new RemoteItem(auto, initImmediate);
      } else if (oleClass.isOlReport()) {
        ret = (E) new ReportItem(auto, initImmediate);
      } else if (oleClass.isOlTask()) {
        ret = (E) new TaskItem(auto, initImmediate);
      } else if (oleClass.isOlTaskRequest()) {
        ret = (E) new TaskRequestAcceptItem(auto, initImmediate);
      } else if (oleClass.isOlTaskRequestDecline()) {
        ret = (E) new TaskRequestDeclineItem(auto, initImmediate);
      } else if (oleClass.isOlTaskRequestUpdate()) {
        ret = (E) new TaskRequestUpdateItem(auto, initImmediate);
      }
    }
    return ret;
  }
}
