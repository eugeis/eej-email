package ee.email.outlook;

import org.eclipse.swt.ole.win32.OleAutomation;

public interface OleAutoFactory<E extends OleAuto> {

  E createOleAutoObject(OleAutomation auto, boolean initImmediate);
}
