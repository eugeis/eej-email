package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210916(v=office.11).aspx">Explorer</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211785(v=office.11).aspx">Caption</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211796(v=office.11).aspx">CommandBars</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211818(v=office.11).aspx">CurrentFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211822(v=office.11).aspx">CurrentView</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171421(v=office.11).aspx">HTMLDocument</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212160(v=office.11).aspx">Height</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171470(v=office.11).aspx">Left</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171816(v=office.11).aspx">Panes</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171941(v=office.11).aspx">Selection</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220423(v=office.11).aspx">Top</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221849(v=office.11).aspx">Width</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221852(v=office.11).aspx">WindowState</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa219395(v=office.11).aspx">Activate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220086(v=office.11).aspx">DeselectFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220118(v=office.11).aspx">IsFolderSelected</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220119(v=office.11).aspx">IsPaneVisible</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210283(v=office.11).aspx">SelectFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210298(v=office.11).aspx">ShowPane</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa209972(v=office.11).aspx">Activate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa209982(v=office.11).aspx">BeforeFolderSwitch</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171125(v=office.11).aspx">BeforeItemCopy</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171136(v=office.11).aspx">BeforeItemCut</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171142(v=office.11).aspx">BeforeItemPaste</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171149(v=office.11).aspx">BeforeMaximize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171155(v=office.11).aspx">BeforeMinimize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171178(v=office.11).aspx">BeforeMove</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171204(v=office.11).aspx">BeforeSize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171207(v=office.11).aspx">BeforeViewSwitch</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171213(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171237(v=office.11).aspx">Deactivate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171252(v=office.11).aspx">FolderSwitch</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171366(v=office.11).aspx">SelectionChange</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa219368(v=office.11).aspx">ViewSwitch</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210948(v=office.11).aspx">MAPIFolder</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210981(v=office.11).aspx">Panes</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211053(v=office.11).aspx">Selection</a>
 *      </p>*
 * @author eugeis
 */

public class Explorer extends OleAuto {

  protected String caption;

  protected Variant commandBars;

  protected Variant currentFolder;

  protected Variant currentView;

  protected Variant hTMLDocument;

  protected Variant height;

  protected Variant left;

  protected Panes panes;

  protected Selection selection;

  protected Variant top;

  protected Variant width;

  protected OlWindowStateEnum windowState;

  public Explorer(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getCaption();
    getCommandBars();
    getCurrentFolder();
    getCurrentView();
    getHTMLDocument();
    getHeight();
    getLeft();
    getPanes();
    getSelection();
    getTop();
    getWidth();
    getWindowState();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211785(v=office.11).aspx">Caption</a>
   */
  public String getCaption() {

    String propertyName = "Caption";
    try {
      if (this.caption == null) {
        this.caption = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.caption;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211796(v=office.11).aspx">CommandBars</a>
   */
  public Variant getCommandBars() {

    String propertyName = "CommandBars";
    try {
      if (this.commandBars == null) {
        this.commandBars = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.commandBars;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211818(v=office.11).aspx">CurrentFolder</a>
   */
  public Variant getCurrentFolder() {

    String propertyName = "CurrentFolder";
    try {
      if (this.currentFolder == null) {
        this.currentFolder = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentFolder;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211822(v=office.11).aspx">CurrentView</a>
   */
  public Variant getCurrentView() {

    String propertyName = "CurrentView";
    try {
      if (this.currentView == null) {
        this.currentView = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentView;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171421(v=office.11).aspx">HTMLDocument</a>
   */
  public Variant getHTMLDocument() {

    String propertyName = "HTMLDocument";
    try {
      if (this.hTMLDocument == null) {
        this.hTMLDocument = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hTMLDocument;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa212160(v=office.11).aspx">Height</a>
   */
  public Variant getHeight() {

    String propertyName = "Height";
    try {
      if (this.height == null) {
        this.height = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.height;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171470(v=office.11).aspx">Left</a>
   */
  public Variant getLeft() {

    String propertyName = "Left";
    try {
      if (this.left == null) {
        this.left = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.left;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171816(v=office.11).aspx">Panes</a>
   */
  public Panes getPanes() {

    String propertyName = "Panes";
    try {
      if (this.panes == null) {
        this.panes = new Panes(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.panes;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171941(v=office.11).aspx">Selection</a>
   */
  public Selection getSelection() {

    String propertyName = "Selection";
    try {
      if (this.selection == null) {
        this.selection = new Selection(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.selection;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220423(v=office.11).aspx">Top</a>
   */
  public Variant getTop() {

    String propertyName = "Top";
    try {
      if (this.top == null) {
        this.top = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.top;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221849(v=office.11).aspx">Width</a>
   */
  public Variant getWidth() {

    String propertyName = "Width";
    try {
      if (this.width == null) {
        this.width = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.width;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221852(v=office.11).aspx">WindowState</a>
   */
  public OlWindowStateEnum getWindowState() {

    String propertyName = "WindowState";
    try {
      if (this.windowState == null) {
        this.windowState = OlWindowStateEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.windowState;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.panes != null) {
      this.panes.dispose();
    }
    if (this.selection != null) {
      this.selection.dispose();
    }
  }

}
