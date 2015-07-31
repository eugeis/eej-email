package ee.email.outlook.base;

import java.util.Date;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211012(v=office.11).aspx">RecurrencePattern</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211826(v=office.11).aspx">DayOfMonth</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211827(v=office.11).aspx">DayOfWeekMask</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211847(v=office.11).aspx">Duration</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211863(v=office.11).aspx">EndTime</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211866(v=office.11).aspx">Exceptions</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171438(v=office.11).aspx">Instance</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171441(v=office.11).aspx">Interval</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171686(v=office.11).aspx">MonthOfYear</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171770(v=office.11).aspx">NoEndDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171772(v=office.11).aspx">Occurrences</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171828(v=office.11).aspx">PatternEndDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171831(v=office.11).aspx">PatternStartDate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171884(v=office.11).aspx">RecurrenceType</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171890(v=office.11).aspx">Regenerate</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172055(v=office.11).aspx">StartTime</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa220110(v=office.11).aspx">GetOccurrence</a>
 *      </p>
 *      <p>
 *      Child Objects | <a href="http://msdn.microsoft.com/en-us/library/aa210911(v=office.11).aspx">Exceptions</a>
 *      </p>*
 * @author eugeis
 */

public class RecurrencePattern extends OleAuto {

  protected Variant dayOfMonth;

  protected OlDaysOfWeekEnum dayOfWeekMask;

  protected Variant duration;

  protected Date endTime;

  protected OleExceptions exceptions;

  protected Variant instance;

  protected Variant interval;

  protected Variant monthOfYear;

  protected Boolean noEndDate;

  protected Variant occurrences;

  protected Date patternEndDate;

  protected Date patternStartDate;

  protected OlRecurrenceTypeEnum recurrenceType;

  protected Boolean regenerate;

  protected Date startTime;

  public RecurrencePattern(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  @Override
  public void init() {

    super.init();
    getDayOfMonth();
    getDayOfWeekMask();
    getDuration();
    getEndTime();
    getExceptions();
    getInstance();
    getInterval();
    getMonthOfYear();
    getNoEndDate();
    getOccurrences();
    getPatternEndDate();
    getPatternStartDate();
    getRecurrenceType();
    getRegenerate();
    getStartTime();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211826(v=office.11).aspx">DayOfMonth</a>
   */
  public Variant getDayOfMonth() {

    String propertyName = "DayOfMonth";
    try {
      if (this.dayOfMonth == null) {
        this.dayOfMonth = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.dayOfMonth;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211827(v=office.11).aspx">DayOfWeekMask</a>
   */
  public OlDaysOfWeekEnum getDayOfWeekMask() {

    String propertyName = "DayOfWeekMask";
    try {
      if (this.dayOfWeekMask == null) {
        this.dayOfWeekMask = OlDaysOfWeekEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.dayOfWeekMask;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211847(v=office.11).aspx">Duration</a>
   */
  public Variant getDuration() {

    String propertyName = "Duration";
    try {
      if (this.duration == null) {
        this.duration = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.duration;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211863(v=office.11).aspx">EndTime</a>
   */
  public Date getEndTime() {

    String propertyName = "EndTime";
    try {
      if (this.endTime == null) {
        this.endTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.endTime;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211866(v=office.11).aspx">Exceptions</a>
   */
  public OleExceptions getExceptions() {

    String propertyName = "Exceptions";
    try {
      if (this.exceptions == null) {
        this.exceptions = new OleExceptions(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.exceptions;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171438(v=office.11).aspx">Instance</a>
   */
  public Variant getInstance() {

    String propertyName = "Instance";
    try {
      if (this.instance == null) {
        this.instance = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.instance;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171441(v=office.11).aspx">Interval</a>
   */
  public Variant getInterval() {

    String propertyName = "Interval";
    try {
      if (this.interval == null) {
        this.interval = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.interval;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171686(v=office.11).aspx">MonthOfYear</a>
   */
  public Variant getMonthOfYear() {

    String propertyName = "MonthOfYear";
    try {
      if (this.monthOfYear == null) {
        this.monthOfYear = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.monthOfYear;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171770(v=office.11).aspx">NoEndDate</a>
   */
  public Boolean getNoEndDate() {

    String propertyName = "NoEndDate";
    try {
      if (this.noEndDate == null) {
        this.noEndDate = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.noEndDate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171772(v=office.11).aspx">Occurrences</a>
   */
  public Variant getOccurrences() {

    String propertyName = "Occurrences";
    try {
      if (this.occurrences == null) {
        this.occurrences = getProperty(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.occurrences;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171828(v=office.11).aspx">PatternEndDate</a>
   */
  public Date getPatternEndDate() {

    String propertyName = "PatternEndDate";
    try {
      if (this.patternEndDate == null) {
        this.patternEndDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.patternEndDate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171831(v=office.11).aspx">PatternStartDate</a>
   */
  public Date getPatternStartDate() {

    String propertyName = "PatternStartDate";
    try {
      if (this.patternStartDate == null) {
        this.patternStartDate = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.patternStartDate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171884(v=office.11).aspx">RecurrenceType</a>
   */
  public OlRecurrenceTypeEnum getRecurrenceType() {

    String propertyName = "RecurrenceType";
    try {
      if (this.recurrenceType == null) {
        this.recurrenceType = OlRecurrenceTypeEnum.findEnum(getIntegerValue(propertyName));
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.recurrenceType;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171890(v=office.11).aspx">Regenerate</a>
   */
  public Boolean getRegenerate() {

    String propertyName = "Regenerate";
    try {
      if (this.regenerate == null) {
        this.regenerate = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.regenerate;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa172055(v=office.11).aspx">StartTime</a>
   */
  public Date getStartTime() {

    String propertyName = "StartTime";
    try {
      if (this.startTime == null) {
        this.startTime = getDateValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.startTime;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.exceptions != null) {
      this.exceptions.dispose();
    }
  }

}
