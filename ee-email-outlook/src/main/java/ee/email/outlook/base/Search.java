package ee.email.outlook.base;

import org.eclipse.swt.ole.win32.OleAutomation;

import ee.email.outlook.OleAuto;

/**
 * @see <p>
 *      Type <a href="http://msdn.microsoft.com/en-us/library/aa211048(v=office.11).aspx">Search</a>
 *      </p>
 *      <p>
 *      Properties | <a href="http://msdn.microsoft.com/en-us/library/aa211393(v=office.11).aspx">Application</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211793(v=office.11).aspx">Class</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa211871(v=office.11).aspx">Filter</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171451(v=office.11).aspx">IsSynchronous</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171818(v=office.11).aspx">Parent</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171925(v=office.11).aspx">Results</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171934(v=office.11).aspx">Scope</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa171938(v=office.11).aspx">SearchSubFolders</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa172019(v=office.11).aspx">Session</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa220278(v=office.11).aspx">Tag</a>
 *      </p>
 *      <p>
 *      Methods | <a href="http://msdn.microsoft.com/en-us/library/aa210281(v=office.11).aspx">Save</a> | <a
 *      href="http://msdn.microsoft.com/en-us/library/aa210314(v=office.11).aspx">Stop</a>
 *      </p>*
 * @author eugeis
 */

public class Search extends OleAuto {

  protected String filter;

  protected Boolean isSynchronous;

  protected Results results;

  protected String scope;

  protected Boolean searchSubFolders;

  protected String tag;

  public Search(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
  }

  public void init() {

    super.init();
    getFilter();
    getIsSynchronous();
    getResults();
    getScope();
    getSearchSubFolders();
    getTag();
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa211871(v=office.11).aspx">Filter</a>
   */
  public String getFilter() {

    String propertyName = "Filter";
    try {
      if (this.filter == null) {
        this.filter = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.filter;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171451(v=office.11).aspx">IsSynchronous</a>
   */
  public Boolean getIsSynchronous() {

    String propertyName = "IsSynchronous";
    try {
      if (this.isSynchronous == null) {
        this.isSynchronous = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.isSynchronous;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171925(v=office.11).aspx">Results</a>
   */
  public Results getResults() {

    String propertyName = "Results";
    try {
      if (this.results == null) {
        this.results = new Results(getPropertyAs(propertyName), initImmediate);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.results;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171934(v=office.11).aspx">Scope</a>
   */
  public String getScope() {

    String propertyName = "Scope";
    try {
      if (this.scope == null) {
        this.scope = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.scope;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa171938(v=office.11).aspx">SearchSubFolders</a>
   */
  public Boolean getSearchSubFolders() {

    String propertyName = "SearchSubFolders";
    try {
      if (this.searchSubFolders == null) {
        this.searchSubFolders = getBooleanValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.searchSubFolders;
  }

  /**
   * @see <a href="http://msdn.microsoft.com/en-us/library/aa220278(v=office.11).aspx">Tag</a>
   */
  public String getTag() {

    String propertyName = "Tag";
    try {
      if (this.tag == null) {
        this.tag = getStringValue(propertyName);
      }
    } catch (Exception e) {
      handleGetPropertyException(e, propertyName);
    }
    return this.tag;
  }

  @Override
  public void dispose() {

    super.dispose();
    if (this.results != null) {
      this.results.dispose();
    }
  }

}
