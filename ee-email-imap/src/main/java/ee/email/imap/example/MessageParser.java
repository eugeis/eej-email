package ee.email.imap.example;

import java.io.IOException;
import java.util.Arrays;

import javax.mail.Address;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;

import com.sun.mail.imap.IMAPMessage;

import ee.email.model.BodyFormatEnum;
import ee.email.model.Email;

public class MessageParser {
  public Email parseMessage(IMAPMessage message) throws MessagingException,
      IOException {
    Email email = new Email();
    email.setId(message.getMessageID());
    email.setFrom(toString(message.getFrom()));
    String contentType = message.getContentType();
    BodyFormatEnum bodyFormat = parseBodyFormat(contentType);
    email.setBodyFormat(bodyFormat);
    if (contentType.contains("multipart")) {
      StringBuffer sb = new StringBuffer();
      extractPart(message, sb);
      email.setBody(sb.toString());
    } else {
      email.setBody(message.getContent().toString());
    }

    email.setTo(toString(message.getRecipients(Message.RecipientType.TO)));
    email.setCc(toString(message.getRecipients(Message.RecipientType.CC)));
    email.setBcc(toString(message.getRecipients(Message.RecipientType.BCC)));
    email.setDate(message.getReceivedDate());
    email.setSubject(message.getSubject());
    return email;
  }

  protected String toString(Address[] addresses) {
    return addresses != null ? Arrays.toString(addresses) : null;
  }

  public BodyFormatEnum parseBodyFormat(String contentType) {
    BodyFormatEnum ret = null;
    String format = contentType.toString().toLowerCase();
    if (format.startsWith("text/html")) {
      ret = BodyFormatEnum.HTML;
    } else if (format.startsWith("text/plain")) {
      ret = BodyFormatEnum.Plain;
    }
    return ret;
  }

  private void extractPart(final Part part, StringBuffer body)
      throws MessagingException, IOException {
    if (part.getContent() instanceof Multipart) {
      Multipart mp = (Multipart) part.getContent();
      for (int i = 0; i < mp.getCount(); i++) {
        extractPart(mp.getBodyPart(i), body);
      }
    } else if (part.getContentType().toLowerCase().startsWith("text/html")) {
      if (body.length() != 0) {
        body.append("<HR/>");
      }
      body.append(part.getContent());
    }
  }
}
