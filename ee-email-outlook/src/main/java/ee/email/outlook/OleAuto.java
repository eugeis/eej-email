package ee.email.outlook;

import java.util.Date;

import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.email.outlook.base.OlObjectClassEnum;

public class OleAuto {

  private static DateConverter theDateConveter = new DateConverter();

  protected final static String TAB = ",";

  private final Logger logger = LoggerFactory.getLogger(OleAuto.class);

  protected OleAutomation auto;

  protected Application application;

  protected OlObjectClassEnum classOle;

  protected Variant parent;

  protected NameSpace session;

  protected boolean initImmediate;

  public OleAuto() {

    super();
  }

  public OleAuto(OleAutomation auto, boolean initImmediate) {

    super();
    this.auto = auto;
    this.initImmediate = initImmediate;

    if (this.initImmediate) {
      init();
    }
  }

  /**
   * Checks if is auto.
   * 
   * @return true, if is auto
   */
  public boolean isAuto() {

    return this.auto != null;
  }

  public void init() {

  }

  protected OleAutomation getAuto() {

    return this.auto;
  }

  protected void setAuto(OleAutomation auto) {

    this.auto = auto;
  }

  protected OleAutomation invokeAsAuto(int[] dispId, int value) {

    return invokeAsAuto(dispId, new Variant(value));
  }

  protected OleAutomation invokeAsAuto(int[] dispId, Variant value) {

    OleAutomation ret = null;
    try {
      ret = getAsAuto(this.auto.invoke(dispId[0], new Variant[] { value }));
    } catch (Exception e) {
      this.logger.error("Exception by invoke of '{}' with '{}' in '{}", new Object[] { dispId, value, this.auto });
    }
    return ret;
  }

  public OleAutomation getPropertyAs(String name) {

    return getAsAuto(getProperty(getId(name)));
  }

  public OleAutomation getAsAuto(Variant varResult) {

    if (varResult != null && varResult.getType() != OLE.VT_EMPTY) {
      OleAutomation result = varResult.getAutomation();
      varResult.dispose();
      return result;
    }
    return null;
  }

  public Variant getProperty(String name) {

    Variant ret = getProperty(getId(name));
    return ret;
  }

  public String[] getStringValues(String... names) {

    String[] ret = new String[names.length];
    for (int i = 0; i < names.length; i++) {
      Variant value = getProperty(getId(names[i]));
      ret[i] = value != null ? value.getString() : null;
    }
    return ret;
  }

  public String getStringValue(String name) {

    String ret = null;
    Variant value = getProperty(getId(name));
    ret = value != null ? value.getString() : null;
    return ret;
  }

  public Integer getIntegerValue(String name) {

    Integer ret = null;
    try {
      Variant value = getProperty(getId(name));
      ret = value != null ? value.getInt() : null;
    } catch (Exception e) {
      this.logger.error("Exception by getProperty of '{}' in '{}", new Object[] { name, this.auto });
    }
    return ret;
  }

  public Boolean getBooleanValue(String name) {

    Boolean ret = null;
    try {
      Variant value = getProperty(getId(name));
      ret = value != null ? value.getBoolean() : null;
    } catch (Exception e) {
      this.logger.error("Exception by getProperty of '{}' in '{}", new Object[] { name, this.auto });
    }
    return ret;
  }

  protected Variant getProperty(Integer id) {

    Variant ret = null;
    if (id != null) {
      ret = this.auto.getProperty(id);
    }
    return ret;
  }

  public Date getDateValue(String name) {

    Date ret = null;
    try {
      Variant value = getProperty(getId(name));
      ret = value != null ? theDateConveter.convertToDate(value) : null;
    } catch (Exception e) {
      this.logger.error("Exception by getProperty of '{}' in '{}", new Object[] { name, this.auto });
    }
    return ret;
  }

  public String getPropertyAsString(String name) {

    String ret = getProperty(name).getString();
    return ret;
  }

  public Variant invoke(String command, String value) {

    return invoke(command, new Variant(value));
  }

  public Variant invoke(String command, Variant... values) {

    Variant ret = null;
    try {
      ret = this.auto.invoke(getId(command), values);
    } catch (Exception e) {
      this.logger.error("Exception by invoke of '{}' with '{}' in '{}", new Object[] { command, values, this.auto });
    }
    return ret;
  }

  public void invokeNoReply(String command, Variant... values) {

    try {
      this.auto.invokeNoReply(getId(command), values);
    } catch (Exception e) {
      this.logger.error("Exception by invoke of '{}' with '{}' in '{}", new Object[] { command, values, this.auto });
    }
  }

  public OleAutomation invokeAsAuto(String command, String value) {

    return invokeAsAuto(command, new Variant(value));
  }

  public OleAutomation invokeAsAuto(String command, int value) {

    return invokeAsAuto(command, new Variant(value));
  }

  public OleAutomation invokeAsAuto(String command, Variant value) {

    OleAutomation ret = null;
    try {
      ret = getAsAuto(invoke(command, value));
    } catch (Exception e) {
      this.logger.error("Exception by invoke of '{}' with '{}' in '{}", new Object[] { command, value, this.auto });
    }
    return ret;
  }

  public Variant invoke(String command) {

    return this.auto.invoke(getId(command));
  }

  public boolean setProperty(String name, String value) {

    if (value != null) {
      return this.auto.setProperty(getId(name), new Variant(value));
    } else {
      return false;
    }
  }

  public boolean setProperty(String name, int value) {

    return this.auto.setProperty(getId(name), new Variant(value));
  }

  private Integer getId(String name) {

    Integer ret = null;
    int[] array = this.auto.getIDsOfNames(new String[] { name });
    if (array != null && array.length >= 1) {
      ret = array[0];
    } else {
      this.logger.error("Id not found for '{}' in '{}", new Object[] { name, this.auto });
    }
    return ret;
  }

  public void dispose() {

    if (this.auto != null) {
      this.auto.dispose();
      this.auto = null;
    }
  }

  public boolean isConnect() {

    return this.auto != null;
  }

  protected void handleGetPropertyException(Exception e, String propertyName) {

    this.logger.error("Exception '{}' by getting of '{}' in '{}", new Object[] { e, propertyName, this.auto });
  }

  public Application getApplication() {

    String propertyName = "Application";
    try {
      if (this.application == null) {
        this.application = new Application(getPropertyAs(propertyName), this.initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.application;
  }

  public OlObjectClassEnum getClassOle() {

    String propertyName = "Class";
    try {
      if (this.classOle == null) {
        this.classOle = OlObjectClassEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.classOle;
  }

  public Variant getParent() {

    String propertyName = "Parent";
    try {
      if (this.parent == null) {
        this.parent = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.parent;
  }

  public NameSpace getSession() {

    String propertyName = "Session";
    try {
      if (this.session == null) {
        this.session = new NameSpace(getPropertyAs(propertyName), this.initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.session;
  }

  /*
   * (non-Javadoc)
   * @see com.siemens.pa.it.siamos.ifc.ofp.domain.OFPbaseObject#toString()
   */
  @Override
  public String toString() {

    StringBuffer buffer = new StringBuffer();
    buffer.append(getClass().getSimpleName());
    buffer.append("@");
    buffer.append(Integer.toHexString(hashCode()));
    buffer.append("[");
    fillToString(buffer);
    buffer.append("]");
    return buffer.toString();
  }

  /**
   * Fill to string.
   * 
   * @param buffer the buffer
   * @return the string buffer
   */
  protected StringBuffer fillToString(StringBuffer buffer) {

    buffer.append("oleClass=").append(getClassOle());
    return buffer;
  }
}
