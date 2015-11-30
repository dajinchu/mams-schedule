package heathers.face;

import java.util.Calendar;
import java.util.Date;

public class Period {
	Date start, end;
	String name;
	static Calendar cal = Calendar.getInstance();
	
	public Period(Date start, Date end, String name){
		this.start = start;
		this.end = end;
		this.name = name;		
	}
	
	public void print(){
		cal.setTime(start);
		System.out.println(cal.toString());
	}
}
