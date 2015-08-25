package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.base.MAPIFolderBase;

/**
 * @see <a href="http://msdn.microsoft.com/en-us/library/aa210948(v=office.11).aspx">MAPIFolder</a>
 * @author eugeis
 */
public class MAPIFolder extends MAPIFolderBase {

  public MAPIFolder(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @SuppressWarnings("unchecked")
  @Override
  public Folders<MAPIFolder> getFolders() {

    if (this.folders == null) {
      this.folders = new Folders<MAPIFolder>(getPropertyAs("Folders"), this.initImmediate, new OleAutoFactory<MAPIFolder>() {

        @Override
        public MAPIFolder createOleAutoObject(OleAutomation auto, boolean initImmediate) {

          return new MAPIFolder(auto, initImmediate);
        }
      });
    }
    return this.folders;
  }
}
