package heathers.face;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.SocketException;
import java.util.ArrayList;

import net.fortuna.ical4j.data.CalendarOutputter;
import net.fortuna.ical4j.model.Calendar;
import net.fortuna.ical4j.model.Date;
import net.fortuna.ical4j.model.DateTime;
import net.fortuna.ical4j.model.Dur;
import net.fortuna.ical4j.model.Property;
import net.fortuna.ical4j.model.TimeZone;
import net.fortuna.ical4j.model.TimeZoneRegistry;
import net.fortuna.ical4j.model.TimeZoneRegistryFactory;
import net.fortuna.ical4j.model.ValidationException;
import net.fortuna.ical4j.model.component.VEvent;
import net.fortuna.ical4j.model.component.VTimeZone;
import net.fortuna.ical4j.model.parameter.TzId;
import net.fortuna.ical4j.model.parameter.Value;
import net.fortuna.ical4j.model.property.CalScale;
import net.fortuna.ical4j.model.property.Name;
import net.fortuna.ical4j.model.property.ProdId;
import net.fortuna.ical4j.model.property.Version;
import net.fortuna.ical4j.model.property.XProperty;
import net.fortuna.ical4j.util.UidGenerator;

public class ICalExporter {

	private Calendar calendar;
	private TimeZone timezone;
	private VTimeZone tz;
	private UidGenerator ug;
	private String name;

	public ICalExporter(String name){
		this.name = name;
		calendar = new Calendar();
		calendar.getProperties().add(new ProdId("-//Events Calendar//iCal4j 1.0//EN"));
		calendar.getProperties().add(Version.VERSION_2_0);
		calendar.getProperties().add(CalScale.GREGORIAN);
		calendar.getProperties().add(new XProperty("X-WR-CALNAME",name));
		TimeZoneRegistry registry = TimeZoneRegistryFactory.getInstance().createRegistry();
		timezone = registry.getTimeZone("America/Mexico_City");
		tz = timezone.getVTimeZone();
	}
	
	public void addEvents(ArrayList<Period> periods){
		/*VTimeZone tz = new VTimeZone();
		TzId tzParam = new TzId(tz.getProperties().getProperty(Property.TZID)
		         .getValue());*/
		try {
			ug = new UidGenerator("mams");
		} catch (SocketException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		java.util.Calendar cal = java.util.Calendar.getInstance();
		cal.set(java.util.Calendar.MONTH, java.util.Calendar.DECEMBER);
		cal.set(java.util.Calendar.DAY_OF_MONTH, 25);

		VEvent christmas = new VEvent(new Date(cal.getTime()), "Christmas Day");
		// initialise as an all-day event..
		christmas.getProperties().add(tz.getTimeZoneId());
		christmas.getProperties().add(ug.generateUid());
		calendar.getComponents().add(christmas);
		for(Period period: periods){
			VEvent e = new VEvent(new DateTime(period.start),new DateTime(period.end),period.name);
			e.getProperties().add(tz.getTimeZoneId());
			e.getProperties().add(ug.generateUid());
			calendar.getComponents().add(e);
			period.print();
			//e.getProperties().getProperty(Property.DTSTART).getParameters().add(tzParam);
		}
		try {
			FileOutputStream fout = new FileOutputStream(name+".ics");
			
			CalendarOutputter outputter = new CalendarOutputter();
			outputter.output(calendar, fout);
			System.out.println("done");
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ValidationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
}
