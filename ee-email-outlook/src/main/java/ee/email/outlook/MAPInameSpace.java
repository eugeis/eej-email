package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

public class MAPInameSpace extends NameSpace {

  private MAPIFolder defaultEmailFolder;

  private MAPIFolder defaultContactFolder;

  public MAPInameSpace(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public MAPIFolder getDefaultEmailFolder() {

    if (this.defaultEmailFolder == null) {
      this.defaultEmailFolder = new MAPIFolder(invokeAsAuto("GetDefaultFolder", 6), this.initImmediate);
    }
    return this.defaultEmailFolder;
  }

  public MAPIFolder getDefaultContactFolder() {

    if (this.defaultContactFolder == null) {
      this.defaultContactFolder = new MAPIFolder(invokeAsAuto("GetDefaultFolder", 10), this.initImmediate);
    }
    return this.defaultContactFolder;
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

  @Override
  public void dispose() {

    super.dispose();
    if (this.defaultEmailFolder != null) {
      this.defaultEmailFolder.dispose();
    }
    if (this.defaultContactFolder != null) {
      this.defaultContactFolder.dispose();
    }
  }
}
