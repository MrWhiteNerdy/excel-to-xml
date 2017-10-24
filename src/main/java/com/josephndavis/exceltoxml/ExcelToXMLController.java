package com.josephndavis.exceltoxml;

import javafx.event.ActionEvent;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;

import java.io.File;

public class ExcelToXMLController {
	
	public Button chooseFileBtn;
	public Button convertBtn;
	public TextField filePathFld;
	public Label msg;
	
	private File inputFile;
	
	public void handleConvertBtnAction(ActionEvent actionEvent) {
		try {
			ExcelToXML.convert(getInputFile());
			msg.setText("File converted successfully");
		} catch (Exception e) {
			msg.setText("Failure: File not converted successfully");
		}
	}
	
	public void handleChooseFileBtnAction(ActionEvent actionEvent) {
		FileChooser fileChooser = new FileChooser();
		File selectedFile = fileChooser.showOpenDialog(null);
		setInputFile(selectedFile);
		fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
		
		FileChooser.ExtensionFilter filter = new FileChooser.ExtensionFilter("Excel files (*.xlsx)", "*.xlsx");
		fileChooser.getExtensionFilters().addAll(filter);
		
		if (selectedFile != null) {
			filePathFld.setText(selectedFile.getName());
			msg.setText("");
		} else {
			msg.setText("Must choose file");
		}
	}
	
	private void setInputFile(File inputFile) {
		this.inputFile = inputFile;
	}
	
	private File getInputFile() {
		return inputFile;
	}
}
