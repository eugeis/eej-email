package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211095(v=office.11).aspx">UserProperties</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220070(v=office.11).aspx">Add</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220093(v=office.11).aspx">Find</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a>
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

public class UserProperties<E extends UserProperty> extends OleCollection<E> {

  public UserProperties(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new OleAutoFactory<E>() {

      @SuppressWarnings("unchecked")
      @Override
      public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

        return (E) new UserProperty(auto, initImmediate);
      }
    });

  }

  public UserProperties(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

    super(auto, initImmediate, childFactory);
  }

  @Override
  public void init() {

    super.init();
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
