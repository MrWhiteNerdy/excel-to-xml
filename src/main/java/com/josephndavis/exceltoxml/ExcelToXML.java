package com.josephndavis.exceltoxml;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelToXML {
	
	/**
	 * Converts an Excel spreadsheet to an XML document. The Excel spreadsheet must have
	 * the first row contain headers for the data. The output file will have the same name
	 * as the input file and will go into the same directory.
	 * @param inputFile The Excel spreadsheet to convert from
	 */
    public static void convert(File inputFile) {
    	// Data structures to hold information from Excel spreadsheet
	    ArrayList<String> headers = new ArrayList<>();
	    ArrayList<String> data = new ArrayList<>();
	    
        try {
        	// Get spreadsheet
            FileInputStream fis = new FileInputStream(inputFile);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
	        
            // Evaluator to convert all values in spreadsheet to strings
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
	
            // Loop through each cell in the spreadsheet to put its value
	        // into either the headers ArrayList or the data ArrayList
	        for (Row row : sheet) {
		        Iterator<Cell> cellIterator = row.cellIterator();
		
		        while (cellIterator.hasNext()) {
			        Cell cell = cellIterator.next();
			
			        if (row.getRowNum() == sheet.getRow(0).getRowNum()) {
				        headers.add(formatCell(cell, evaluator));
			        } else {
				        data.add(formatCell(cell, evaluator));
			        }
		        }
	        }
	
	        // Create the XML document
	        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
	        DocumentBuilder builder = factory.newDocumentBuilder();
	        Document document = builder.newDocument();

	        // Create the root element and append to document
	        Element root = document.createElement("data");
	        document.appendChild(root);
	        
	        Element headerElement;
	        Element dataElement;
	        
	        // Loop through the spreadsheet data and add it all
	        // to the XML document
	        for (int i = 0; i < sheet.getLastRowNum(); i++) {
		        headerElement = document.createElement("row");
		        headerElement.setAttribute("id", Integer.toString(i + 1));
	        	root.appendChild(headerElement);
	        	
	        	for (int j = 0; j < headers.size(); j++) {
			        dataElement = document.createElement(headers.get(j));
			        headerElement.appendChild(dataElement);
			        dataElement.appendChild(document.createTextNode(data.get((headers.size() - 1) * i + (j + i))));
		        }
	        }
	        
	        fis.close();
	
	        File outputFile = determineOutputFile(inputFile);

            // Stream XMl document to output file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            StreamResult result = new StreamResult(outputFile);
            transformer.transform(source, result);
        } catch (IOException | ParserConfigurationException | TransformerException e) {
            e.printStackTrace();
        }
    }

    // Formats each cell value to be a string
    private static String formatCell(Cell cell, FormulaEvaluator evaluator) {
        DataFormatter df = new DataFormatter();
        return df.formatCellValue(evaluator.evaluateInCell(cell));
    }
    
    // Determines the path of the output file based off the path of the input file
    private static File determineOutputFile(File inputFile) {
	    String outputFileName = inputFile.getName().substring(0,
			    inputFile.getName().lastIndexOf(".")) + ".xml";
	    String outputFilePath = inputFile.getAbsolutePath().substring(0,
			    inputFile.getAbsolutePath().lastIndexOf("\\"));
	
	    return new File(outputFilePath + File.separator + outputFileName);
    }
	
}