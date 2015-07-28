package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa210923(v=office.11).aspx">Inspector</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211785(v=office.11).aspx">Caption</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211796(v=office.11).aspx">CommandBars</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211820(v=office.11).aspx">CurrentItem</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211848(v=office.11).aspx">EditorType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171425(v=office.11).aspx">HTMLEditor</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa212160(v=office.11).aspx">Height</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171470(v=office.11).aspx">Left</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171495(v=office.11).aspx">ModifiedFormPages</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220423(v=office.11).aspx">Top</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221849(v=office.11).aspx">Width</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221852(v=office.11).aspx">WindowState</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa221855(v=office.11).aspx">WordEditor</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa219395(v=office.11).aspx">Activate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220077(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220090(v=office.11).aspx">Display</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220117(v=office.11).aspx">HideFormPage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220121(v=office.11).aspx">IsWordMail</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/dd594060(v=office.11).aspx">SetControlItemProperty</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210290(v=office.11).aspx">SetCurrentFormPage</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210296(v=office.11).aspx">ShowFormPage</a>
 *      </p>
 *      <p>
 *      Events | <a href="http://msdn.microsoft.com/en-us/library/aa209972(v=office.11).aspx">Activate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171149(v=office.11).aspx">BeforeMaximize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171155(v=office.11).aspx">BeforeMinimize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171178(v=office.11).aspx">BeforeMove</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171204(v=office.11).aspx">BeforeSize</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171213(v=office.11).aspx">Close</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171237(v=office.11).aspx">Deactivate</a>
 *      </p>*
 * @author eugeis
 */

public class Inspector extends OleAuto {

  protected String caption;

  protected Variant commandBars;

  protected Variant currentItem;

  protected OlEditorTypeEnum editorType;

  protected Variant hTMLEditor;

  protected Variant height;

  protected Variant left;

  protected Variant modifiedFormPages;

  protected Variant top;

  protected Variant width;

  protected OlWindowStateEnum windowState;

  protected Variant wordEditor;

  public Inspector(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getCaption();
    getCommandBars();
    getCurrentItem();
    getEditorType();
    getHTMLEditor();
    getHeight();
    getLeft();
    getModifiedFormPages();
    getTop();
    getWidth();
    getWindowState();
    getWordEditor();
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211820(v=office.11).aspx">CurrentItem</a>
   */
  public Variant getCurrentItem() {

    String propertyName = "CurrentItem";
    try {
      if (this.currentItem == null) {
        this.currentItem = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.currentItem;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211848(v=office.11).aspx">EditorType</a>
   */
  public OlEditorTypeEnum getEditorType() {

    String propertyName = "EditorType";
    try {
      if (this.editorType == null) {
        this.editorType = OlEditorTypeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.editorType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171425(v=office.11).aspx">HTMLEditor</a>
   */
  public Variant getHTMLEditor() {

    String propertyName = "HTMLEditor";
    try {
      if (this.hTMLEditor == null) {
        this.hTMLEditor = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.hTMLEditor;
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
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171495(v=office.11).aspx">ModifiedFormPages</a>
   */
  public Variant getModifiedFormPages() {

    String propertyName = "ModifiedFormPages";
    try {
      if (this.modifiedFormPages == null) {
        this.modifiedFormPages = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.modifiedFormPages;
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

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa221855(v=office.11).aspx">WordEditor</a>
   */
  public Variant getWordEditor() {

    String propertyName = "WordEditor";
    try {
      if (this.wordEditor == null) {
        this.wordEditor = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.wordEditor;
  }

  @Override
  public void dispose() {

    super.dispose();
  }

}
