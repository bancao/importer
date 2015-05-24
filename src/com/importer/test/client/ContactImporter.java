package com.importer.test.client;

import java.io.File;
import java.io.IOException;
import java.util.List;

import com.importer.ExcelImporter;
import com.importer.model.Contact;

public class ContactImporter {
	
	public static void main(String[] args) throws IOException,
			InstantiationException, IllegalAccessException {
		File file = new File("E:\\workspace\\excelImport\\src\\com\\excel\\client\\contacts.xls");
		List<Object> contacts = ExcelImporter.importFile(Contact.class, file);
		for (Object object : contacts) {
			Contact contact = (Contact) object;
			System.out.println(contact.getFirstName());
			System.out.println(contact.getLastName());
		}
	}
	
	

}
