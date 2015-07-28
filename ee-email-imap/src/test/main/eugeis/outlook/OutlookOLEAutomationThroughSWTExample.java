package eugeis.email.imap;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

/**
 * @author Tom
 */
public class OutlookOLEAutomationThroughSWTExample {

  /**
   * @param args
   */
  public static void main(String[] args) {

    Display display = new Display();
    Shell shell = new Shell(display);

    shell.setText("Outlook Automation");
    shell.setLayout(new FillLayout());

    OleFrame frm = new OleFrame(shell, SWT.NONE);

    OleClientSite site = new OleClientSite(frm, SWT.NONE, "Outlook.Application");

    OleAutomation auto = new OleAutomation(site);

    int[] GetNamespaceDispId = auto.getIDsOfNames(new String[] { "GetNamespace" });
    Variant mapiNamespace = auto.invoke(GetNamespaceDispId[0], new Variant[] { new Variant("MAPI") });

    OleAutomation mapiNamespaceAuto = mapiNamespace.getAutomation();

    int[] DefaultFolderPropertyDispId = mapiNamespaceAuto.getIDsOfNames(new String[] { "GetDefaultFolder" });

    // 6 is default for emails (Inbox)
    // 10 is default for contacs
    Variant defaultFolder = mapiNamespaceAuto.invoke(DefaultFolderPropertyDispId[0], new Variant[] { new Variant(10) });

    OleAutomation defaultFolderAutomation = defaultFolder.getAutomation();

    int[] ItemsFolderPropertyDispId = defaultFolderAutomation.getIDsOfNames(new String[] { "Items" });

    Variant items = defaultFolderAutomation.invoke(ItemsFolderPropertyDispId[0]);

    OleAutomation itemsAutomation = items.getAutomation();

    int[] ItemsCountPropertyDispId = itemsAutomation.getIDsOfNames(new String[] { "Count" });

    int[] itemDispId = itemsAutomation.getIDsOfNames(new String[] { "Item" });

    Variant itemsCount = itemsAutomation.invoke(ItemsCountPropertyDispId[0]);

    for (int i = 1, cnt = itemsCount.getInt(); i <= cnt; i++) {
      Variant contact = itemsAutomation.invoke(itemDispId[0], new Variant[] { new Variant(i) });
      OleAutomation contactAutomation = contact.getAutomation();

      int[] ContactNamePropertyDispId = contactAutomation.getIDsOfNames(new String[] { "FullName" });

      Variant contactName = contactAutomation.getProperty(ContactNamePropertyDispId[0]);

      int[] CompanyNamePropertyDispId = contactAutomation.getIDsOfNames(new String[] { "CompanyName" });

      Variant companyName = contactAutomation.getProperty(CompanyNamePropertyDispId[0]);

      System.out.println("Contact: " + contactName.getString() + " works for " + companyName.getString());

      contactAutomation.dispose();
    }

    itemsAutomation.dispose();
    defaultFolderAutomation.dispose();

    mapiNamespaceAuto.dispose();
    shell.dispose();
    auto.dispose();

    site.deactivateInPlaceClient();
    site.dispose();

    frm.dispose();
  }
}