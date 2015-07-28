package ee.email.core.elastic;

import java.util.List;

public interface Indexer<E> {

  void index(String parentReference, E entityToIndex);

  void index(String parentReference, List<E> entityListToIndex);
}