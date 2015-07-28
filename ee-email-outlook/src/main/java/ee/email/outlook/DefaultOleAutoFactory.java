package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

public class DefaultOleAutoFactory<E extends OleAuto> implements OleAutoFactory<E> {

  @SuppressWarnings("unchecked")
  @Override
  public E createOleAutoObject(OleAutomation auto, boolean initImmediate) {

    return (E) new OleAuto(auto, initImmediate);
  }

}
