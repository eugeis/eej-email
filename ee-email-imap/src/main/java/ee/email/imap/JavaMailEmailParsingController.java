/*
 * Copyright 2015-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package ee.email.imap;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Properties;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sun.mail.imap.IMAPFolder;
import com.sun.mail.imap.IMAPMessage;

import ee.email.core.EmailParsingController;
import ee.email.core.FolderFilter;
import ee.email.core.ParsedCallback;
import ee.email.imap.example.MessageParser;
import ee.email.model.Email;

public class JavaMailEmailParsingController implements
    EmailParsingController<Email> {
  private final static Logger logger = LoggerFactory
      .getLogger(JavaMailEmailParsingController.class);

  private String host;
  private String email;
  private String password;
  private String emailprovider;
  private FolderFilter folderFilterForRecursion;
  private FolderFilter folderFilter;
  private MessageParser messageParser;

  public JavaMailEmailParsingController(String host, String email,
      String password, String emailprovider,
      FolderFilter folderFilterForRecursion, FolderFilter folderFilter) {
    super();
    this.host = host;
    this.email = email;
    this.password = password;
    this.emailprovider = emailprovider;
    messageParser = new MessageParser();
    this.folderFilterForRecursion = folderFilterForRecursion;
    this.folderFilter = folderFilter;
  }

  @Override
  public int parseEmailContainer(File file,
      ParsedCallback<Email> parsedcallback, Date date) throws IOException {

    return 0;
  }

  @Override
  public int parseEmails(ParsedCallback<Email> parsedcallback, Date date)
      throws IOException {
    int ret = 0;
    Properties props = System.getProperties();
    props.setProperty("mail.store.protocol", "imaps");
    props.put("mail.imap.fetchsize", "1048576");
    props.setProperty("mail.imap.partialfetch", "false");
    props.setProperty("mail.imaps.partialfetch", "false");
    Store store = null;
    try {
      Session session = Session.getDefaultInstance(props, null);
      store = session.getStore(emailprovider);
      store.connect(host, email, password);
      ExecutorService executor = Executors.newFixedThreadPool(10);
      parseEmails(store.getDefaultFolder(), parsedcallback, date, executor);
      executor.shutdown();
      try {
        executor.awaitTermination(24, TimeUnit.HOURS);
      } catch (InterruptedException e) {
        // nothing
      }
    } catch (MessagingException e) {
      logger.error("Exception '{}' by connection '{}'", e, store);
    } finally {
      try {
        if (store != null) {
          store.close();
        }
      } catch (MessagingException e) {
        logger.error("Exception '{}' by closing fo store '{}'", e, store);
      }
    }
    return ret;
  }

  protected void parseEmails(final Folder folder,
      final ParsedCallback<Email> parsedcallback, Date lastDate,
      ExecutorService executor) {
    try {
      if (folder.exists()
          && (folderFilter == null || folderFilter.isFolderToParse(folder
              .getFullName()))) {
        parseMessages(folder, parsedcallback, lastDate, executor);
      } else {
        logger.info("Ignore messages of the folder '{}'", folder.getFullName());
      }
    } catch (Exception e) {
      logger.error("Unexcpected exception '{}' by email parsing in '{}'", e,
          folder);
    }

    // is folder to parse recursive
    if (folderFilterForRecursion == null
        || folderFilterForRecursion.isFolderToParse(folder.getFullName())) {
      try {
        logger.info("Parse childs of the folder '{}'", folder.getFullName());
        for (Folder childFolder : folder.list()) {
          parseEmails(childFolder, parsedcallback, lastDate, executor);
        }
      } catch (Exception e) {
        // nothing
      }
    } else {
      logger.info("Ignore childs of the folder '{}'", folder.getFullName());
    }
  }

  protected void parseMessages(final Folder folder,
      final ParsedCallback<Email> parsedcallback, Date lastDate,
      ExecutorService executor) throws MessagingException {
    logger.info("Parse messages of the folder '{}'", folder.getFullName());
    if (folder instanceof IMAPFolder) {
      final IMAPFolder imapFolder = (IMAPFolder) folder;
      imapFolder.open(Folder.READ_ONLY);
      final String imapFolderName = imapFolder.getFullName();
      if (folder.isOpen()) {
        Message[] messages = imapFolder.getMessages();

        if (lastDate == null) {
          for (int i = 0; i < messages.length; i++) {
            final Message message = messages[i];
            parseEmail(message, imapFolderName, parsedcallback, executor);
          }
        } else {
          for (int i = messages.length - 1; i >= 0; i--) {
            final Message message = messages[i];
            if (message.getReceivedDate().after(lastDate)) {
              parseEmail(message, imapFolderName, parsedcallback, executor);
            } else {
              // all new messages read
              break;
            }
          }
        }

      } else {
        logger.error("Folder '{}' is still closed", imapFolderName);
      }
    }
  }

  private void parseEmail(final Message message, final String folderName,
      final ParsedCallback<Email> parsedcallback, ExecutorService executor) {
    try {
      final Email email = messageParser.parseMessage((IMAPMessage) message);
      executor.execute(new Runnable() {
        @Override
        public void run() {
          try {
            parsedcallback.parsed(folderName, email);
          } catch (Exception e) {
            logger.error("'{}' by email callback processing in '{}'", e,
                folderName);
          }
        }
      });
    } catch (Exception e) {
      logger.error("'{}' by email parsing in '{}'", e, folderName);
    }
  }
}
