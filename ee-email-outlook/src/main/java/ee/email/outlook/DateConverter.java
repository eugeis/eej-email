package ee.email.outlook;

import java.util.Calendar;
import java.util.Date;
import org.eclipse.swt.ole.win32.Variant;

public class DateConverter {

  private final Calendar calendar = Calendar.getInstance();

  public Date convertToDate(Variant dateAsVariant) {

    this.calendar.set(1899, 11, 30, 0, 0, 0);
    float julianDate = dateAsVariant.getFloat();
    int days = (int) julianDate;
    float minutes = julianDate - days;
    this.calendar.add(Calendar.DATE, days);
    this.calendar.add(Calendar.MINUTE, (int) ((60 * 24) * minutes));
    return this.calendar.getTime();
  }
}
