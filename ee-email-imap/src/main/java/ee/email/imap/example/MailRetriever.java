package ee.email.imap.example;

import java.io.IOException;
import java.util.Properties;

import javax.mail.FetchProfile;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Store;

public class MailRetriever {

  private String emailuser;
  private String emailpassword;
  private String emailserver;
  private String emailprovider;

  public MailRetriever(String emailuser, String emailpassword, String emailserver, String emailprovider) {
    this.emailuser = emailuser;
    this.emailpassword = emailpassword;
    this.emailserver = emailserver;
    this.emailprovider = emailprovider;
  }

  public void getMail() {
    Session session;
    Store store = null;
    Folder folder = null;
    Folder inboxfolder = null;

    Properties props = System.getProperties();
    props.setProperty("mail.pop3s.rsetbeforequit", "true");
    props.setProperty("mail.pop3.rsetbeforequit", "true");
    session = Session.getInstance(props, null);
    //     session.setDebug(true);

    try {
      store = session.getStore(emailprovider);
      store.connect(emailserver, emailuser, emailpassword);
      folder = store.getDefaultFolder();
      if (folder == null)
        throw new Exception("No default folder");
      inboxfolder = folder.getFolder("INBOX");
      if (inboxfolder == null)
        throw new Exception("No INBOX");
      inboxfolder.open(Folder.READ_ONLY);

      Message[] msgs = inboxfolder.getMessages();

      FetchProfile fp = new FetchProfile();
      fp.add("Subject");
      inboxfolder.fetch(msgs, fp);

      for (int j = msgs.length - 1; j >= 0; j--) {
        if (msgs[j].getSubject().startsWith("MailPage:")) {
          setLatestMessage(msgs[j]);
          break;
        }
      }

      inboxfolder.close(false);
      store.close();

    } catch (NoSuchProviderException ex) {
      ex.printStackTrace();
    } catch (MessagingException ex) {
      ex.printStackTrace();
    } catch (Exception ex) {
      ex.printStackTrace();
    } finally {
      try {
        if (store != null)
          store.close();
      } catch (MessagingException ex) {
        ex.printStackTrace();
      }
    }
  }

  public Renderable getLatestMessage() {
    return latestMessage;
  }

  private Renderable latestMessage;

  void setLatestMessage(Message message) {
    if (message == null) {
      latestMessage = null;
      return;
    }

    try {
      if (message.getContentType().startsWith("text/plain")) {
        latestMessage = new RenderablePlainText(message);
      } else {
        latestMessage = new RenderableMessage(message);
      }
    } catch (MessagingException ex) {
      ex.printStackTrace();
    } catch (IOException ex) {
      ex.printStackTrace();
    }
  }

  public static void main(String[] args) {
    MailRetriever mr = new MailRetriever(args[0], args[1], args[2], args[3]);
    mr.getMail();
    Renderable msg = mr.getLatestMessage();
    if (msg == null) {
      System.out.println("No valid messages in the mail account");
    } else {
      System.out.println("Subject:" + msg.getSubject());
      System.out.println("Body Text:" + msg.getBodytext());
      System.out.println(msg.getAttachmentCount() + " attachments");
      for (int i = 0; i < msg.getAttachmentCount(); i++) {
        Attachment at = msg.getAttachment(i);
        System.out.println(at.getFilename() + " " + at.getContent().length + " bytes of (" + at.getContenttype() + ")");
      }
    }
  }
}
