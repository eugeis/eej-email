/*
 * RenderablePlainText.java
 *
 * Created on 10 November 2005, 10:49
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 */

package ee.email.imap.example;

import java.io.IOException;

import javax.mail.Message;
import javax.mail.MessagingException;

public class RenderablePlainText implements Renderable {
    
    String bodytext;
    String subject;
    
    public RenderablePlainText(Message message) throws MessagingException, IOException {
        subject=message.getSubject().substring("MailPage:".length());
        bodytext=(String)message.getContent();
    }
    
    public Attachment getAttachment(int i) {
        return null;
    }
    
    public int getAttachmentCount() {
        return 0;
    }
    
    public String getBodytext() {
        return "<PRE>"+bodytext+"</PRE>";
    }
    
    public String getSubject() {
        return subject;
    }
    
}
