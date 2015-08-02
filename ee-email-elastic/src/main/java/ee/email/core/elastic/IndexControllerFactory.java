package ee.email.core.elastic;

import java.io.IOException;
import java.util.Date;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ee.elastic.ElasticAdmin;
import ee.elastic.IndexAdmin;
import ee.elastic.Mapping;
import ee.elastic.NodeType;
import ee.email.core.EmailParsingController;
import ee.email.core.EmailParsingFactory;
import ee.email.model.Email;

/**
 * The Class IndexControllerFactory.
 *
 * @author Eugen Eisler
 * @lastChangedBy $Author:$
 * @date $Date:$
 * @version $Revision:$
 */
public class IndexControllerFactory {

  private static Logger logger = LoggerFactory
      .getLogger(IndexControllerFactory.class);

  private static final String lastEmailsDateProperty = "lastEmailDate";

  private static final String lastEmailsIdProperty = "lastEmailsId";

  private final String indexName;

  private final EmailParsingFactory<Email> emailParsingFactory;

  public IndexControllerFactory(String indexName,
      EmailParsingFactory<Email> emailParsingFactory) {

    this.indexName = indexName;
    this.emailParsingFactory = emailParsingFactory;
  }

  public void recreateIndex(NodeType nodeType) {

    ElasticAdmin elasticAdmin = buildElasticAdmin(nodeType);
    IndexAdmin indexAdmin = new IndexAdmin(indexName, elasticAdmin);
    EmailParsingController<Email> emailParsingController = emailParsingFactory
        .getEmailParsingController();
    IndexerOfParsedEmails parsedCallback = new IndexerOfParsedEmails(
        new IndexerImpl(indexAdmin));

    try {
      indexAdmin.connect();
      recreateIndex(indexAdmin, parsedCallback, emailParsingController);
    } finally {
      dispose(indexAdmin);
    }
  }

  protected ElasticAdmin buildElasticAdmin(NodeType nodeType) {
    ElasticAdmin elasticAdmin = new ElasticAdmin(nodeType);
    if (nodeType == NodeType.Transport) {
      elasticAdmin.addServerAddress("localhost", 9300);
    }
    return elasticAdmin;
  }

  protected void recreateIndex(IndexAdmin indexAdmin,
      IndexerOfParsedEmails parsedCallback,
      EmailParsingController<Email> emailParsingController) {

    indexAdmin.recreateIndex(buildMappings());
    try {
      emailParsingController.parseEmails(parsedCallback, null);
      storeIndexProperties(indexAdmin, parsedCallback);
    } catch (Exception e) {
      logger.error("Exception {} by indexing of {}", e, emailParsingFactory);
    }
  }

  public void synchronizeIndex(NodeType nodeType) {

    ElasticAdmin elasticAdmin = buildElasticAdmin(nodeType);
    IndexAdmin indexAdmin = new IndexAdmin(indexName, elasticAdmin);
    EmailParsingController<Email> emailParsingController = emailParsingFactory
        .getEmailParsingController();
    IndexerOfParsedEmails parsedCallback = new IndexerOfParsedEmails(
        new IndexerImpl(indexAdmin));
    try {
      indexAdmin.connect();
      if (indexAdmin.checkIndex()) {
        synchronizeIndex(indexAdmin, parsedCallback, emailParsingController);
      } else {
        recreateIndex(indexAdmin, parsedCallback, emailParsingController);
      }
    } finally {
      dispose(indexAdmin);
    }
  }

  protected void dispose(IndexAdmin indexAdmin) {
    if (indexAdmin != null) {
      indexAdmin.close();
    }
  }

  protected void synchronizeIndex(IndexAdmin indexAdmin,
      IndexerOfParsedEmails parsedCallback,
      EmailParsingController<Email> emailParsingController) {

    try {
      Date lastEmailDate = indexAdmin.propertyAsDate(lastEmailsDateProperty);
      Integer lastIdProperty = (Integer) indexAdmin
          .property(lastEmailsIdProperty);
      parsedCallback.setMaxId(lastIdProperty != null ? lastIdProperty : 0);
      emailParsingController.parseEmails(parsedCallback, lastEmailDate);
      storeIndexProperties(indexAdmin, parsedCallback);
    } catch (Exception e) {
      logger.error("Exception {} by indexing of outlook", e);
    }
  }

  protected void storeIndexProperties(IndexAdmin indexAdmin,
      IndexerOfParsedEmails parsedCallback) throws IOException {
    storeIndexProperty(indexAdmin, lastEmailsDateProperty,
        parsedCallback.getMaxEmailDate());
    storeIndexProperty(indexAdmin, lastEmailsIdProperty,
        parsedCallback.getMaxId());
  }

  protected void storeIndexProperty(IndexAdmin indexAdmin, String name,
      Object value) throws IOException {

    if (value != null) {
      indexAdmin.property(name, value);
    }
  }

  @SuppressWarnings("unchecked")
  public static void main(String[] args) {

    if (args != null && args.length == 3) {
      EmailParsingFactory<Email> emailParsingFactory = null;
      try {
        String command = args[0].trim();
        String indexName = args[1].trim();
        String classOfEmailParsingFactory = args[2].trim();

        emailParsingFactory = (EmailParsingFactory<Email>) Class.forName(
            classOfEmailParsingFactory).newInstance();
        IndexControllerFactory indexControllerFactory = new IndexControllerFactory(
            indexName, emailParsingFactory);
        if (command.equalsIgnoreCase("recreate")) {
          indexControllerFactory.recreateIndex(NodeType.Transport);
        } else if (command.equalsIgnoreCase("synchronize")) {
          indexControllerFactory.synchronizeIndex(NodeType.Transport);
        } else if (command.equalsIgnoreCase("updateMappings")) {
          indexControllerFactory.updateMappings(NodeType.Transport);
        } else {
          indexControllerFactory.synchronizeIndex(NodeType.Transport);
        }
      } catch (Exception e) {
        e.printStackTrace();
        logger
            .error(
                "Exception {} occured in IndexControllerFactory, with parameters {}: {}",
                e, args, e);
      } finally {
        if (emailParsingFactory != null) {
          try {
            emailParsingFactory.close();
          } catch (IOException e) {
            logger.error("Exception {} by closing of emailParsingFactory", e,
                args);
          }
        }
      }
    } else {
      printHelp();
    }
  }

  public void updateMappings(NodeType nodeType) {

    IndexAdmin indexAdmin = new IndexAdmin(indexName,
        new ElasticAdmin(nodeType));

    try {
      indexAdmin.connect();
      indexAdmin.createMappings(buildMappings());
    } finally {
      dispose(indexAdmin);
    }

  }

  protected Mapping[] buildMappings() {
    return new Mapping[] { new Mapping("email",
        "/ee/email/elasticsearch/email-mapping.json") };
  }

  protected static void printHelp() {

    System.out
        .println("IndexControllerFactory [command] [indexName] [classOfEmailParsingFactory]");
    System.out.println("Command:");
    System.out.println("\trecreate");
    System.out.println("\tsynchronize");
    System.out.println("\tupdateMappings");
    System.out.println();
  }
}
