package ee.email.outlook;

import java.util.Iterator;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

public class OleCollection<E extends OleAuto> extends OleAuto implements Iterable<E> {

  private int[] itemDispId = null;

  protected Variant count;

  protected OleAutoFactory<E> itemFactory;

  public OleCollection(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> itemFactory) {

    super(auto, initImmediate);
    this.itemFactory = itemFactory;
  }

  @Override
  public void init() {

    this.count = getProperty("Count");
  }

  private void loadItemNameId() {

    if (this.itemDispId == null) {
      this.itemDispId = this.auto.getIDsOfNames(new String[] { "Item" });
    }
  }

  public E getItem(int number) {

    loadItemNameId();
    return this.itemFactory.createOleAutoObject(invokeAsAuto(this.itemDispId, number), this.initImmediate);
  }

  public Variant getCount() {

    if (this.count == null) {
      this.count = getProperty("Count");
    }
    return this.count;
  }

  @Override
  public Iterator<E> iterator() {

    return new Iterator<E>() {

      private int number = 0;

      private final int count = getCount().getInt();

      @Override
      public boolean hasNext() {

        return this.count > this.number;
      }

      @Override
      public E next() {

        return getItem(++this.number);
      }

      @Override
      public void remove() {

        throw new RuntimeException("The method is not supported");
      }
    };
  }

}
