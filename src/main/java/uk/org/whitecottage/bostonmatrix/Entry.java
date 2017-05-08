package uk.org.whitecottage.bostonmatrix;

import static uk.org.whitecottage.poi.SheetHelper.getCellNumericValue;
import static uk.org.whitecottage.poi.SheetHelper.getCellStringValue;

import org.apache.poi.ss.usermodel.Row;

public class Entry implements Comparable<Entry> {
	protected String name = "";
	protected String key = "";
	protected double x;
	protected double y;
	protected double z;
	
	public Entry(Row row) {
		if (row != null) {
			key = getCellStringValue(row, 1);
			name = getCellStringValue(row, 2);
			x = getCellNumericValue(row, 3);
			y = getCellNumericValue(row, 4);
			z = getCellNumericValue(row, 5);
		}
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public double getX() {
		return x;
	}

	public void setX(double x) {
		this.x = x;
	}

	public double getY() {
		return y;
	}

	public void setY(double y) {
		this.y = y;
	}

	public double getZ() {
		return z;
	}

	public void setZ(double z) {
		this.z = z;
	}

	@Override
	public int compareTo(Entry o) {
		return (int) (z - o.getZ());
	}
	
	@Override
	public boolean equals(Object o) {
		if (o == null) {
			return false;
		}
		
		if (!(o instanceof Entry)) {
			return false;
		}
		
		return z == ((Entry) o).getZ();
	}
}