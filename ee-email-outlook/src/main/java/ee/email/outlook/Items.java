package ee.email.outlook;

import java.util.Iterator;

import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.Variant;

import ee.email.outlook.base.ItemsBase;

/**
 * @see <a href="http://msdn.microsoft.com/en-us/library/aa210932(v=office.11).aspx">Items</a>
 * @author eugeis
 */
public class Items<E extends OleAuto> extends ItemsBase implements Iterable<E> {

  private int[] itemDispId = null;

  protected OleAutoFactory<E> itemFactory;

  @SuppressWarnings({ "unchecked", "rawtypes" })
  public Items(OleAutomation auto, boolean initImmediate) {

    super(auto, initImmediate);
    this.itemFactory = new ItemOleAutoFactory();
  }

  public Items(OleAutomation auto, boolean initImmediate, OleAutoFactory<E> itemFactory) {

    super(auto, initImmediate);
    this.itemFactory = itemFactory;
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

  /**
   * Sort.
   * 
   * @param property the property
   * @param descending the descending
   */
  public void sort(String property, boolean descending) {

    invokeNoReply("Sort", new Variant("[" + property + "]"), new Variant(descending));
  }
}
