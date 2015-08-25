package ee.email.core.elastic;

import static org.elasticsearch.common.xcontent.XContentFactory.jsonBuilder;

import java.util.List;

import org.elasticsearch.action.index.IndexRequestBuilder;
import org.elasticsearch.common.xcontent.XContentBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.elastic.IndexAdmin;
import ee.email.model.Email;

public class IndexerImpl implements Indexer<Email> {

  public static final String _SOURCE = "_source";

  public static final String _ALL = "_all";

  public static final String EMAIL = "email";

  public static final String SUBJECT = "subject";

  public static final String FOLDER = "folder";

  public static final String FROM = "from";

  public static final String FROM_NAME = "fromName";

  public static final String TO = "to";

  public static final String BODY = "body";

  public static final String BODY_FORMAT = "bodyFormat";

  public static final String DATE = "date";

  public Logger logger = LoggerFactory.getLogger(IndexerImpl.class);

  private final IndexAdmin indexAdmin;

  public IndexerImpl(IndexAdmin indexAdmin) {

    super();
    this.indexAdmin = indexAdmin;
  }

  @Override
  public void index(String parentReference, List<Email> items) {

    for (Email item : items) {
      index(parentReference, item);
    }
  }

  @Override
  public void index(String parentReference, Email item) {

    // on startup
    try {
      this.logger.info("indexing of {}", item);

      // add class info
      XContentBuilder builder = jsonBuilder().startObject();
      builder.field(FOLDER, parentReference);
      builder.field(SUBJECT, item.getSubject() != null ? item.getSubject() : "");
      builder.field(FROM, item.getFrom() != null ? item.getFrom() : "");
      builder.field(BODY, item.getBody() != null ? item.getBody() : "");
      builder.field(BODY_FORMAT, item.getBodyFormat() != null ? item.getBodyFormat().name() : "");
      builder.field(DATE, item.getDate() != null ? item.getDate() : "");
      builder.field(FROM_NAME, item.getFromName() != null ? item.getFromName() : "");
      builder.field(TO, item.getTo() != null ? item.getTo() : "");

      IndexRequestBuilder requestBuilder = this.indexAdmin.prepareIndex(EMAIL, item.getId()).setSource(builder.endObject());

      // IndexResponse response = requestBuilder.execute().actionGet();
      requestBuilder.execute();

    } catch (Throwable e) {
      this.logger.error("exception {} at index of {}", e, item);
    }
  }
}
