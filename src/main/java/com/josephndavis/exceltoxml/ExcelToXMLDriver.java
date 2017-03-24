package com.josephndavis.exceltoxml;

import java.io.File;

public class ExcelToXMLDriver {
	
	public static void main(String[] args) {
		File file = new File(args[0]);
		
		ExcelToXML.convert(file);
	}
	
}
