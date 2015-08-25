package ee.email.imap;

import java.io.IOException;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.email.core.EmailParsingController;
import ee.email.core.ParsedCallback;
import ee.email.core.RegExpFolderFilter;
import ee.email.model.Email;

public class EmailParsingFactory implements
    ee.email.core.EmailParsingFactory<Email> {

  private final Logger logger = LoggerFactory
      .getLogger(EmailParsingFactory.class);

  /**
   * e.g. '.*(Inbox|Sent Items).*'
   */
  public final static String SYS__REG_EXP_FOR_FOLDER = "regExpFolder";

  /**
   * e.g. '.*(<LastName>, <FirstName>|2010|2011).*'
   */
  public final static String SYS__REG_EXP_FOR_FOLDER_RECURSION = "regExpFolderRecursion";

  public final static String SYS__SERVER = "server";

  public final static String SYS__LOGIN = "login";

  public final static String SYS__PASSWORD = "password";

  @Override
  public void close() throws IOException {

  }

  @Override
  public EmailParsingController<Email> getEmailParsingController() {
    return new JavaMailEmailParsingController(
        findRequiredSystemProperty(SYS__SERVER),
        findRequiredSystemProperty(SYS__LOGIN), findRequiredSystemProperty(
            SYS__PASSWORD, true), "imaps", createFolderFilterForRecursion(),
        createFolderFilter());
  }

  /**
   * Creates the folder filter.
   * 
   * @return the reg exp folder filter
   */
  protected RegExpFolderFilter createFolderFilter() {

    RegExpFolderFilter folderFilter = new RegExpFolderFilter(
        findRequiredSystemProperty(SYS__REG_EXP_FOR_FOLDER, ".*", false), true);
    return folderFilter;
  }

  /**
   * Creates the folder filter for recursion.
   * 
   * @return the reg exp folder filter
   */
  protected RegExpFolderFilter createFolderFilterForRecursion() {

    RegExpFolderFilter folderFilterForRecursion = new RegExpFolderFilter(
        findRequiredSystemProperty(SYS__REG_EXP_FOR_FOLDER_RECURSION, ".*",
            false), true);
    return folderFilterForRecursion;
  }

  private String findRequiredSystemProperty(String key) {
    return findRequiredSystemProperty(key, null, false);
  }

  private String findRequiredSystemProperty(String key, boolean hidden) {
    return findRequiredSystemProperty(key, null, hidden);
  }

  private String findRequiredSystemProperty(String key, String defaultValue,
      boolean hidden) {
    String ret = System.getProperty(key, defaultValue);
    if (ret == null) {
      throw new IllegalArgumentException("System parameter '" + key
          + "' not defined.");
    } else {
      this.logger.info("Use system parameter {}={}", key, hidden ? "..." : ret);
    }
    return ret;
  }

  public static void main(String[] args) throws IOException {
    new EmailParsingFactory().getEmailParsingController().parseEmails(
        new ParsedCallback<Email>() {

          @Override
          public void parsed(String parentReference, List<Email> entities) {
            System.out.println(entities);
          }

          @Override
          public void parsed(String parentReference, Email entity) {
            System.out.println(entity);
          }
        }, null);
  }

}
