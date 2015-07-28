package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210897(v=office.11).aspx">Application</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211387(v=office.11).aspx">AnswerWizard</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211405(v=office.11).aspx">Assistant</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211795(v=office.11).aspx">COMAddIns</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211868(v=office.11).aspx">Explorers</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171437(v=office.11).aspx">Inspectors</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171459(v=office.11).aspx">LanguageSettings</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171852(v=office.11).aspx">ProductCode</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171902(v=office.11).aspx">Reminders</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221802(v=office.11).aspx">Version</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa219397(v=office.11).aspx">ActiveExplorer</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219400(v=office.11).aspx">ActiveInspector</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219402(v=office.11).aspx">ActiveWindow</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220071(v=office.11).aspx">AdvancedSearch</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220078(v=office.11).aspx">CopyFile</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220082(v=office.11).aspx">CreateItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220081(v=office.11).aspx">CreateItemFromTemplate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220083(v=office.11).aspx">CreateObject</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220108(v=office.11).aspx">GetNameSpace</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220120(v=office.11).aspx">IsSearchSynchronous</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220133(v=office.11).aspx">Quit</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa209973(v=office.11).aspx">AdvancedSearchComplete</a>
 *      | <a href="http://msdn.microsoft.com/en-us/library/aa209974(v=office.11).aspx">AdvancedSearchStopped</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171284(v=office.11).aspx">ItemSend</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171290(v=office.11).aspx">MapiLogonComplete</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171307(v=office.11).aspx">NewMail</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171304(v=office.11).aspx">NewMailEx</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171319(v=office.11).aspx">OptionsPagesAdd</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171328(v=office.11).aspx">Quit</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171363(v=office.11).aspx">Reminder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219363(v=office.11).aspx">Startup</a>
 *      </p>*
 * @author eugeis
 */

public class ApplicationBase extends OleAuto {

  protected Variant answerWizard;

  protected Variant assistant;

  protected Variant cOMAddIns;

  protected Explorers explorers;

  protected Inspectors inspectors;

  protected Variant languageSettings;

  protected Variant name;

  protected String productCode;

  protected Reminders reminders;

  protected String version;

  public ApplicationBase(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getAnswerWizard();
    getAssistant();
    getCOMAddIns();
    getExplorers();
    getInspectors();
    getLanguageSettings();
    getName();
    getProductCode();
    getReminders();
    getVersion();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211387(v=office.11).aspx">AnswerWizard</a>
   */
  public Variant getAnswerWizard() {

    String propertyName = "AnswerWizard";
    try {
      if (this.answerWizard == null) {
        this.answerWizard = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.answerWizard;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211405(v=office.11).aspx">Assistant</a>
   */
  public Variant getAssistant() {

    String propertyName = "Assistant";
    try {
      if (this.assistant == null) {
        this.assistant = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.assistant;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211795(v=office.11).aspx">COMAddIns</a>
   */
  public Variant getCOMAddIns() {

    String propertyName = "COMAddIns";
    try {
      if (this.cOMAddIns == null) {
        this.cOMAddIns = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.cOMAddIns;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211868(v=office.11).aspx">Explorers</a>
   */
  public Explorers getExplorers() {

    String propertyName = "Explorers";
    try {
      if (this.explorers == null) {
        this.explorers = new Explorers(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.explorers;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171437(v=office.11).aspx">Inspectors</a>
   */
  public Inspectors getInspectors() {

    String propertyName = "Inspectors";
    try {
      if (this.inspectors == null) {
        this.inspectors = new Inspectors(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.inspectors;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171459(v=office.11).aspx">LanguageSettings</a>
   */
  public Variant getLanguageSettings() {

    String propertyName = "LanguageSettings";
    try {
      if (this.languageSettings == null) {
        this.languageSettings = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.languageSettings;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171690(v=office.11).aspx">Name</a>
   */
  public Variant getName() {

    String propertyName = "Name";
    try {
      if (this.name == null) {
        this.name = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.name;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171852(v=office.11).aspx">ProductCode</a>
   */
  public String getProductCode() {

    String propertyName = "ProductCode";
    try {
      if (this.productCode == null) {
        this.productCode = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.productCode;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171902(v=office.11).aspx">Reminders</a>
   */
  public Reminders getReminders() {

    String propertyName = "Reminders";
    try {
      if (this.reminders == null) {
        this.reminders = new Reminders(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.reminders;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221802(v=office.11).aspx">Version</a>
   */
  public String getVersion() {

    String propertyName = "Version";
    try {
      if (this.version == null) {
        this.version = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.version;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.explorers != null) {
      this.explorers.dispose();
    }
    if (this.inspectors != null) {
      this.inspectors.dispose();
    }
    if (this.reminders != null) {
      this.reminders.dispose();
    }
  }

}
