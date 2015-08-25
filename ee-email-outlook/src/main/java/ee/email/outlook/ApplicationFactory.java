package ee.email.outlook;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.email.core.EmailParsingController;
import ee.email.core.EmailParsingFactory;
import ee.email.core.RegExpFolderFilter;
import ee.email.model.Email;

public class ApplicationFactory implements EmailParsingFactory<Email> {
  private final Logger logger = LoggerFactory.getLogger(ApplicationFactory.class);

  /**
   * e.g. '.*(<LastName>, <FirstName>|2010|2011).*'
   */
  public final static String SYS__REG_EXP_FOR_FOLDER_RECURSION = "regExpFolderRecursion";

  /**
   * e.g. '.*(Inbox|Sent Items).*'
   */
  public final static String SYS__REG_EXP_FOR_FOLDER = "regExpFolder";

  private Display display;

  private Application application;

  private EmailParsingController<Email> emailParsingController;

  public synchronized Application getApplication(boolean initImmediate) {

    if (application == null) {
      Display display = new Display();
      Shell shell = new Shell(display);

      shell.setText("Outlook Automation");
      shell.setLayout(new FillLayout());

      // Open or 'activate' Outlook
      OleFrame frm = new OleFrame(shell, SWT.NONE);
      // This should start outlook if it is not running yet
      //OleClientSite site0 = new OleClientSite(frm, SWT.NONE, "OVCtl.OVCtl");
      //site0.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
      // Now get the outlook application
      OleClientSite site = new OleClientSite(frm, SWT.NONE, "Outlook.Application");
      OleAutomation auto = new OleAutomation(site);

      application = new Application(auto, initImmediate);
      this.display = display;

    }
    return application;
  }

  /*
   * (non-Javadoc)
   *
   * @see ee.email.core.EmailParsingFactory#getEmailParsingController()
   */
  @Override
  public EmailParsingController<Email> getEmailParsingController() {

    if (emailParsingController == null) {
      RegExpFolderFilter folderFilterForRecursion = createFolderFilterForRecursion();
      RegExpFolderFilter folderFilter = createFolderFilter();
      emailParsingController = new OutlookParsingController(getApplication(false), folderFilterForRecursion, folderFilter);
      ;
    }
    return emailParsingController;
  }

  /**
   * Creates the folder filter.
   *
   * @return the reg exp folder filter
   */
  protected RegExpFolderFilter createFolderFilter() {

    RegExpFolderFilter folderFilter = new RegExpFolderFilter(getRequiredSystemProperty(SYS__REG_EXP_FOR_FOLDER), true);
    return folderFilter;
  }

  /**
   * Creates the folder filter for recursion.
   *
   * @return the reg exp folder filter
   */
  protected RegExpFolderFilter createFolderFilterForRecursion() {

    RegExpFolderFilter folderFilterForRecursion = new RegExpFolderFilter(getRequiredSystemProperty(SYS__REG_EXP_FOR_FOLDER_RECURSION), true);
    return folderFilterForRecursion;
  }

  @Override
  public void close() {

    if (application != null) {
      application.dispose();
    }
    if (display != null) {
      display.dispose();
      display = null;
    }
  }

  private String getRequiredSystemProperty(String key) {
    String ret = System.getProperty(key);
    if (ret == null) {
      throw new IllegalArgumentException("System parameter '" + key + "' not defined.");
    } else {
      logger.info("Use system parameter {}={}", key, ret);
    }
    return ret;
  }
}
