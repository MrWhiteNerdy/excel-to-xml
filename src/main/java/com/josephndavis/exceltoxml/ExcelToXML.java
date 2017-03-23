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
	
	private static ArrayList<String> headers = new ArrayList<>();
    private static ArrayList<String> rolesList = new ArrayList<>();
    private static ArrayList<String> data = new ArrayList<>();
	
	/**
	 * Converts an Excel spreadsheet to an XML document. The Excel spreadsheet must have
	 * the first row contain headers for the data. The output file will have the same name
	 * as the input file and will go into the same directory.
	 * @param inputFile The Excel spreadsheet to convert from
	 */
    public static void convert(File inputFile) {
        try {
            FileInputStream fis = new FileInputStream(inputFile);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);
            
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document document = builder.newDocument();
			
	        Element root = document.createElement("company");
            document.appendChild(root);
            Element element;
	        
            Iterator<Row> rowIterator = wb.getSheetAt(0).iterator();

            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                
            Row firstRow = wb.getSheetAt(0).getRow(0);
              
            int firstCol = -1;
            int columnArray[] = new int[firstRow.getLastCellNum()];
                
            for (int i = 0; i < columnArray.length; i++) {
                Cell cell = firstRow.getCell(i);
                String text = cell.getStringCellValue().toLowerCase();
                switch (text) {
                    case "role":
                        firstCol = i;
                        break;
                    case "position":
                        columnArray[0] = i;
                        break;
                    case "id":
                        columnArray[1] = i;
                        break;
                    case "firstname":
                        columnArray[2] = i;
                        break;
                    case "lastname":
                        columnArray[3] = i;
                        break;
                    case "experience":
                        columnArray[4] = i;
                        break;
                    case "salary":
                        columnArray[5] = i;
                        break;
                    case "unit":
                        columnArray[6] = i;
                        break;
                }
            }

            // go through every cell of sheet
            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                
                // if first row, put those values into headera arraylist
                if (nextRow.getRowNum() == firstRow.getRowNum()) {
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        headers.add(formatCell(cell, evaluator));
                    }
                // put all the other values in data arraylist
                // except for values in first column
                } else {
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        if (cell.getColumnIndex() != firstCol) {
                            data.add(formatCell(cell, evaluator));
                        }
                    }
                }
            }
            
            // go through sheet and get the values from first column,
            // excluding header, and put into rolesList arraylist
            for (Row row : sheet) {
                if (row.getRowNum() == firstRow.getRowNum())
                    continue;
                Cell entityCell = row.getCell(firstCol);
                rolesList.add(formatCell(entityCell, evaluator));
            }
            
            // this is where values are put into xml document
            // go through rolesList arraylist and make each value an element and append to root element
            for (int i = 0; i < rolesList.size(); i++) {
                element = document.createElement(rolesList.get(i).toLowerCase());
                root.appendChild(element);
                // for each role, go through other headers to make them an attribute of element
                // also make elements in data arrayList attribute values
                for (int j = 0; j < headers.size() - 1; j++) {
                    // based in which role is being created, 
                    // make appropriate element in data arrayList its attribute value
                    if (i == 0) {
                        element.setAttribute(headers.get(columnArray[j]).toLowerCase(), data.get(j));
                    }
                    if (i == 1) {
                        element.setAttribute(headers.get(columnArray[j]).toLowerCase(), data.get(j + 7));
                    }
                    if (i == 2) {
                        element.setAttribute(headers.get(columnArray[j]).toLowerCase(), data.get(j + 14));
                    }
                    if (i == 3) {
                        element.setAttribute(headers.get(columnArray[j]).toLowerCase(), data.get(j + 21));
                    }
                }
            }
            
            // close file
            fis.close();
                
            // transform xml document to display as file
            // and stream to file location
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            DOMSource source = new DOMSource(document);
            
            File outputFile = determineOutputFile(inputFile);
	        
            StreamResult result = new StreamResult(outputFile);
                
            transformer.transform(source, result);
                
        } catch (IOException | ParserConfigurationException | TransformerException e) {
            e.printStackTrace();
        }
    }

    // this method parses data in each cell to act as string data type
    private static String formatCell(final Cell cell, final FormulaEvaluator evaluator) {
        DataFormatter df = new DataFormatter();
        return df.formatCellValue(evaluator.evaluateInCell(cell));
    }
    
    private static File determineOutputFile(File inputFile) {
	    String outputFileName = inputFile.getName().substring(0, inputFile.getName().lastIndexOf(".")) + ".xml";
	    String outputFilePath = inputFile.getAbsolutePath().substring(0, inputFile.getAbsolutePath().lastIndexOf("\\"));
	
	    return new File(outputFilePath + File.separator + outputFileName);
    }
    
    public static void main(String[] args) {
    	File inputFile = new File("test\\haha\\employees.xlsx");
    	
    	convert(inputFile);
    }
}