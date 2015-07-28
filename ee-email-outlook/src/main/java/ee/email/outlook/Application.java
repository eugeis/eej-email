package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.base.ApplicationBase;

/**
 * @see <a href="http://msdn.microsoft.com/en-us/library/aa210897(v=office.11).aspx">Application</a>
 * @author eugeis
 */
public class Application extends ApplicationBase {

  private MAPInameSpace mapiNamespace;

  public Application(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public MAPInameSpace getMapiNamespace() {

    if (this.mapiNamespace == null) {
      this.mapiNamespace = new MAPInameSpace(invokeAsAuto("GetNamespace", "MAPI"), this.initImmediate);
    }
    return this.mapiNamespace;
  }

  @Override
  public void dispose() {

    super.dispose();

    if (this.mapiNamespace != null) {
      this.mapiNamespace.dispose();
    }
  }
}
