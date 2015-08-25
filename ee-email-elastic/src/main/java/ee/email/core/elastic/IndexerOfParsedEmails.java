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
package ee.email.core.elastic;

import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

import ee.elastic.IndexAdmin;
import ee.email.core.ParsedCallback;
import ee.email.model.Email;

public class IndexerOfParsedEmails implements ParsedCallback<Email> {
  public static final String lastEmailsDateProperty = "lastEmailDate";
  public static final String lastEmailsIdProperty = "lastEmailsId";

  protected IndexAdmin indexAdmin;

  private final Indexer<Email> indexer;

  protected Date maxEmailDate;

  protected AtomicLong maxId;

  public IndexerOfParsedEmails(Indexer<Email> indexer, IndexAdmin indexAdmin) {

    super();
    this.indexAdmin = indexAdmin;
    this.indexer = indexer;
    maxId = new AtomicLong(1);
  }

  public void setMaxId(long maxId) {
    this.maxId.set(maxId + 1);
  }

  @Override
  public void parsed(String parentReference, Email entity) {
    fillGenericIdIfNotSet(entity);
    indexer.index(parentReference, entity);
    checkMaxEmailDate(entity);
    storeIndexProperties();
  }

  protected void checkMaxEmailDate(Email entity) {

    Date emailDate = entity.getDate();
    if (maxEmailDate == null
        || (emailDate != null && emailDate.after(maxEmailDate))) {
      maxEmailDate = emailDate;
    }
  }

  @Override
  public void parsed(String parentReference, List<Email> entities) {
    for (Email entity : entities) {
      fillGenericIdIfNotSet(entity);
    }
    indexer.index(parentReference, entities);
    for (Email email : entities) {
      checkMaxEmailDate(email);
    }

    storeIndexProperties();
  }

  private void fillGenericIdIfNotSet(Email entity) {
    if (entity.getId() == null || entity.getId().isEmpty()) {
      entity.setId(Long.toString(maxId.getAndIncrement()));
    }
  }

  public Date getMaxEmailDate() {
    if (maxEmailDate == null) {
      maxEmailDate = getMaxEmailDateInIndex();
    }
    return maxEmailDate;
  }

  public Date getMaxEmailDateInIndex() {
    return indexAdmin.propertyAsDate(lastEmailsDateProperty);
  }

  public long getMaxId() {
    if (maxId == null) {
      Integer lastIndexedId = getMaxIdInIndex();
      setMaxId(lastIndexedId != null ? lastIndexedId : 0);
    }
    return maxId.get() - 1;
  }

  public Integer getMaxIdInIndex() {
    return (Integer) indexAdmin.property(lastEmailsIdProperty);
  }

  public void storeIndexProperties() {
    storeIndexProperty(indexAdmin, lastEmailsDateProperty, maxEmailDate);
    storeIndexProperty(indexAdmin, lastEmailsIdProperty, maxId.get() - 0);
  }

  protected void storeIndexProperty(IndexAdmin indexAdmin, String name,
    Object value) {

    if (value != null) {
      try {
        indexAdmin.property(name, value);
      } catch (Exception e) {
        e.printStackTrace();
      }
    }
  }
}
