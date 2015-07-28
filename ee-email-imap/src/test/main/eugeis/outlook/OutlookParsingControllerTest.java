package eugeis.email.imap;

import static org.junit.Assert.fail;
import java.io.IOException;
import java.util.Calendar;
import java.util.List;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import eugeis.email.core.EmailParsingController;
import eugeis.email.core.ParsedCallback;
import eugeis.email.core.RegExpFolderFilter;
import eugeis.email.domain.Email;
import eugeis.email.imap.Application;
import eugeis.email.imap.ApplicationFactory;
import eugeis.email.imap.OutlookParsingController;

public class OutlookParsingControllerTest {

  @BeforeClass
  public static void setUpBeforeClass() throws Exception {

  }

  @AfterClass
  public static void tearDownAfterClass() throws Exception {

  }

  // @Test
  public void testParseEmailContainerFileParsedCallback() {

    fail("Not yet implemented");
  }

  // @Test
  public void testParseEmailContainerFile() {

    fail("Not yet implemented");
  }

  @Test
  public void testParseEmail() throws IOException {

    ApplicationFactory applicationFactory = new ApplicationFactory();
    Application application = applicationFactory.createApplication(false);

    EmailParsingController<Email> parsingController = new OutlookParsingController(application);
    Calendar calendar = Calendar.getInstance();
    calendar.set(2011, 07, 18);
    RegExpFolderFilter folderFilterForRecursion = new RegExpFolderFilter(".*(Eisler, Eugen|Wichtig|2011).*", true);
    RegExpFolderFilter folderFilter = new RegExpFolderFilter(".*(Inbox|Sent Items|WICHTIG|Tools|Links).*", true);
    parsingController.parseEmails(new ParsedCallback<Email>() {

      @Override
      public void parsed(String parentReference, Email entity) {

        System.out.println(entity);

      }

      @Override
      public void parsed(String parentReference, List<Email> entities) {

        for (Email entity : entities) {
          System.out.println(entity);
        }
      }

    }, folderFilterForRecursion, folderFilter, calendar.getTime());

    applicationFactory.dispose();

  }

  // @Test
  public void testParseContainer() {

    fail("Not yet implemented");
  }

}
