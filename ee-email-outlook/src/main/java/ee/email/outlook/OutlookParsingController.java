package ee.email.outlook;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.email.core.EmailParsingController;
import ee.email.core.FolderFilter;
import ee.email.core.ParsedCallback;
import ee.email.model.BodyFormatEnum;
import ee.email.model.Email;
import ee.email.outlook.base.MailItem;
import ee.email.outlook.base.OlBodyFormatEnum;
import ee.email.outlook.base.OlItemTypeEnum;

public class OutlookParsingController implements EmailParsingController<Email> {
  private final Logger logger = LoggerFactory.getLogger(OleAuto.class);
  private final RichTextToHtml richTextToHtml = new RichTextToHtml();
  private final Application outlook;

  private FolderFilter folderFilterForRecursion;
  private FolderFilter folderFilter;

  public OutlookParsingController(Application outlookApplication, FolderFilter folderFilterForRecursion, FolderFilter folderFilter) {

    super();
    outlook = outlookApplication;
    this.folderFilterForRecursion = folderFilterForRecursion;
    this.folderFilter = folderFilter;

  }

  @Override
  public int parseEmailContainer(File file, ParsedCallback<Email> parsedCallback,

    Date newerAsDate) throws IOException {

    return 0;
  }

  private int parseFolders(Folders<MAPIFolder> folders, ParsedCallback<Email> parsedCallback, Date newerAsDate) {

    int ret = 0;
    if (folders != null && folders.isAuto()) {
      for (MAPIFolder folder : folders) {
        ret += parseFolder(folder, parsedCallback, newerAsDate);
      }
    }
    return ret;
  }

  private int parseFolder(MAPIFolder folder, ParsedCallback<Email> parsedCallback, Date newerAsDate) {

    int ret = 0;

    String folderName = folder.getFolderPath();
    if (folderName != null) {

      // shall we parse this folder
      OlItemTypeEnum itemType = folder.getDefaultItemType();
      if ((itemType != null && itemType.isOlMailItem()) && (folderFilter == null || folderFilter.isFolderToParse(folderName))) {
        ret += doParseFolder(folderName, folder, parsedCallback, newerAsDate);
      } else {
        logger.info("Ignore messages of the folder '{}'", folderName);
      }

      // shall we parse child?
      if ((folderFilterForRecursion == null || folderFilterForRecursion.isFolderToParse(folderName))) {
        ret += parseFolders(folder.getFolders(), parsedCallback, newerAsDate);
      } else {
        logger.info("Ignore childs of the folder '{}'", folderName);
      }
    }

    folder.dispose();
    return ret;
  }

  protected int doParseFolder(String folderName, MAPIFolder folder, ParsedCallback<Email> parsedCallback, Date newerAsDate) {

    int ret = 0;
    logger.info("Parse folder {} {}", folderName, folder);
    try {
      @SuppressWarnings("unchecked")
      Items<OleAuto> items = folder.getItems();
      if (items != null && items.isAuto()) {
        if (newerAsDate == null) {

          // parse all messages
          for (OleAuto item : items) {
            if (item instanceof MailItem) {
              Email email = parseEmail((MailItem) item);
              if (email != null) {
                parsedCallback.parsed(folderName, email);
                ret++;
              }
            }
            item.dispose();
          }
        } else {
          // parse only new messages
          items.sort("ReceivedTime", true);
          if (items != null && items.isAuto()) {
            for (OleAuto item : items) {
              if (item instanceof MailItem) {
                MailItem mailItem = (MailItem) item;
                Date receivedTime = mailItem.getReceivedTime();
                if (receivedTime != null && receivedTime.before(newerAsDate)) {
                  // no no newer messages
                  break;
                } else {
                  Email email = parseEmail(mailItem);
                  if (email != null) {
                    parsedCallback.parsed(folderName, email);
                    ret++;
                  }
                }
              }
              item.dispose();
            }
          }
        }
      }
    } catch (Exception e) {
      logger.error("Exception {} by parsing of folder", e, folder);
    }
    return ret;
  }

  protected Email parseEmail(MailItem item) {

    Email ret = new Email();
    try {
      ret.setFrom(item.getSenderEmailAddress());
      ret.setSubject(item.getSubject());
      ret.setText(item.getBody());
      ret.setDate(item.getReceivedTime());
      ret.setFromName(item.getSenderName());
      ret.setTo(item.getTo());
      ret.setId(item.getEntryID());

      OlBodyFormatEnum bodyFormat = item.getBodyFormat();
      if (bodyFormat != null) {
        ret.setBodyFormat(BodyFormatEnum.findEnum(bodyFormat.getValue()));
        if (bodyFormat.isOlFormatHTML()) {
          ret.setBody(item.getHTMLBody());
        } else if (bodyFormat.isOlFormatRichText()) {
          ret.setBody(richTextToHtml.rtfToHtml(item.getBody()));
        } else {
          ret.setBody(item.getBody());
        }
      }
    } catch (Exception e) {
      logger.error("Exception {} by parsing of email {}", e, item);
    }
    return ret;
  }

  @Override
  public int parseEmails(ParsedCallback<Email> parsedCallback, Date newerAsDate) throws IOException {

    int ret = 0;

    MAPInameSpace mapi = null;
    try {
      mapi = outlook.getMapiNamespace();
      Folders<MAPIFolder> folders = mapi.getFolders();
      ret = +parseFolders(folders, parsedCallback, newerAsDate);
    } finally {
      if (mapi != null) {
        mapi.dispose();
      }
    }
    return ret;
  }
}
