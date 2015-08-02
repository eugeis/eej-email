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
package ee.email.model;

import java.text.SimpleDateFormat;
import java.util.Date;

public class Email {

  private static SimpleDateFormat dateFormat = new SimpleDateFormat();

  private String id;

  private Date date;

  private String fromName;

  private String from;

  private String to;

  private String cc;

  private String bcc;

  private BodyFormatEnum bodyFormat;

  private String body;

  private String text;

  private String subject;

  public Date getDate() {

    return this.date;
  }

  public void setDate(Date date) {

    this.date = date;
  }

  public String getFromName() {

    return this.fromName;
  }

  public void setFromName(String fromName) {

    this.fromName = fromName;
  }

  public String getFrom() {

    return this.from;
  }

  public void setFrom(String from) {

    this.from = from;
  }

  public String getTo() {

    return this.to;
  }

  public void setTo(String to) {

    this.to = to;
  }

  public String getCc() {

    return this.cc;
  }

  public void setCc(String cc) {

    this.cc = cc;
  }

  public String getBcc() {

    return this.bcc;
  }

  public void setBcc(String bcc) {

    this.bcc = bcc;
  }

  public BodyFormatEnum getBodyFormat() {

    return this.bodyFormat;
  }

  public void setBodyFormat(BodyFormatEnum bodyFormat) {

    this.bodyFormat = bodyFormat;
  }

  public String getBody() {

    return this.body;
  }

  public void setBody(String body) {

    this.body = body;
  }

  public String getText() {

    return this.text;
  }

  public void setText(String text) {

    this.text = text;
  }

  public String getSubject() {

    return this.subject;
  }

  public void setSubject(String subject) {

    this.subject = subject;
  }

  public String getId() {

    return this.id;
  }

  public void setId(String id) {

    this.id = id;
  }

  @Override
  public String toString() {

    return "Email [id=" + this.id + ", from=" + this.from + ", subject=" + this.subject + ", to=" + this.to
        + ", date=" + (this.date != null ? dateFormat.format(this.date) : "null") + "]";
  }

}
