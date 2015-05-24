package com.importer.model;

import com.importer.annotation.ExcelCell;

public class Contact {
	@ExcelCell(columnName="First Name")
	private String firstName;
	
	@ExcelCell(columnName="Last Name")
	private String lastName;

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	public String getLastName() {
		return lastName;
	}

	public void setLastName(String lastName) {
		this.lastName = lastName;
	}

	public static Object newInstance() {
		// TODO Auto-generated method stub
		return null;
	}
}
