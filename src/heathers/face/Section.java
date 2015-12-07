package heathers.face;

import java.util.ArrayList;

public class Section {
	ArrayList<Period> periods = new ArrayList<Period>();
	
	public Section merge(Section section){
		this.periods.addAll(section.periods);
		return this;
	}
}
