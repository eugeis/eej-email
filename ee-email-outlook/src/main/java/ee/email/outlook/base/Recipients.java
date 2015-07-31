package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210995(v=office.11).aspx">Recipients</a>
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
 *      href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210269(v=office.11).aspx">ResolveAll</a>
 *      </p>
 *      <p>
 *      Parent Objects | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210899(v=office.11).aspx">AppointmentItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210936(v=office.11).aspx">JournalItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210946(v=office.11).aspx">MailItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210951(v=office.11).aspx">MeetingItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211067(v=office.11).aspx">TaskItem</a>
 *      </p>*
 * @author eugeis
 */

public class Recipients<E extends Recipient> extends OleCollection<E> {

  public Recipients(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new OleAutoFactory<E>() {

      @SuppressWarnings("unchecked")
      @Override
      public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

        return (E) new Recipient(auto, initImmediate);
      }
    });

  }

  public Recipients(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

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
