package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAutoFactory;
import ee.email.outlook.OleCollection;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211016(v=office.11).aspx">Reminders</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211816(v=office.11).aspx">Count</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220122(v=office.11).aspx">Item</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220142(v=office.11).aspx">Remove</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa171188(v=office.11).aspx">BeforeReminderShow</a> |
 *      <a href="http://msdn.microsoft.com/en-us/library/aa171359(v=office.11).aspx">ReminderAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171360(v=office.11).aspx">ReminderChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171361(v=office.11).aspx">ReminderFire</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171362(v=office.11).aspx">ReminderRemove</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219362(v=office.11).aspx">Snooze</a>
 *      </p>*
 * @author eugeis
 */

public class Reminders<E extends Reminder> extends OleCollection<E> {

  public Reminders(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate, new OleAutoFactory<E>() {

      @SuppressWarnings("unchecked")
      @Override
      public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

        return (E) new Reminder(auto, initImmediate);
      }
    });

  }

  public Reminders(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> childFactory) {

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
