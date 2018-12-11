package com.excel.utility;

import java.lang.invoke.MethodHandles;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.excel.utility.writter.TagWritter;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.FontWeight;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class ExcelApplication extends Application {
	TagWritter tagWritter = new TagWritter();
	private static Logger LOGGER = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getSimpleName());

	@Override
	public void start(Stage primaryStage) {

		primaryStage.setTitle("Excel utility");
		GridPane grid = getScene();
		Scene scene = new Scene(grid, 600, 400);
		Text scenetitle = new Text("Welcome");
		scenetitle.setFont(Font.font("Tahoma", FontWeight.NORMAL, 20));
		grid.add(scenetitle, 0, 0, 2, 1);

		Label xslFileName = new Label("XslFile :");
		grid.add(xslFileName, 0, 1);

		TextField userTextField = new TextField();
		grid.add(userTextField, 1, 1);
		userTextField.setText("data1.xlsx");

		TextField inputFile = new TextField();
		inputFile.setText("wordlines.txt");
		grid.add(inputFile, 1, 2);

		Label pw = new Label("Input Txt File:");
		grid.add(pw, 0, 2);

		CheckBox removeDuplicates = new CheckBox("check duplicate lines and remove");
		grid.add(removeDuplicates, 0, 3);

		CheckBox shuffel = new CheckBox("Random shuffel of the lines");
		grid.add(shuffel, 0, 4);

		CheckBox writeToXsl = new CheckBox("Write the lines to excel");
		grid.add(writeToXsl, 0, 5);

		Button btn = new Button("Run");
		HBox hbBtn = new HBox(10);
		hbBtn.setAlignment(Pos.BOTTOM_RIGHT);
		hbBtn.getChildren().add(btn);
		grid.add(hbBtn, 1, 7);

		final Text actiontarget = new Text();
		grid.add(actiontarget, 1, 9);

		btn.setOnAction(new EventHandler<ActionEvent>() {

			@Override
			public void handle(ActionEvent e) {
				actiontarget.setFill(Color.FIREBRICK);
				String xslPath = userTextField.getText().toString();
				if (xslPath == null || xslPath.isEmpty()) {
					xslPath = "data1.xlsx";
				}
				String inputFile2 = inputFile.getText();
				if (inputFile2 == null || inputFile2.isEmpty()) {
					inputFile2 = "wordlines.txt";
				}
				boolean duplicates = removeDuplicates.isSelected();
				boolean isShuffel = shuffel.isSelected();
				boolean writeExcel = writeToXsl.isSelected();
				String msg = "Files generated sucess fully";
				try {
					if (writeExcel) {
						tagWritter.writetoExcel(xslPath, inputFile2);
					}
					if (isShuffel || duplicates) {
						tagWritter.removeDuplicateAndShuffle(inputFile2, isShuffel, duplicates);
					} else {
						if (!writeExcel) {
							msg = "Select atleast one action";
						}
					}
					actiontarget.setText(msg);

				} catch (Exception e2) {
					LOGGER.error("Error while processing {}", e2);
					actiontarget.setText(e2.getMessage());
				}

			}
		});
		primaryStage.setScene(scene);
		primaryStage.show();
	}

	private GridPane getScene() {
		GridPane grid = new GridPane();
		grid.setAlignment(Pos.CENTER);
		grid.setHgap(10);
		grid.setVgap(10);
		grid.setPadding(new Insets(25, 25, 25, 25));
		return grid;
	}

	public static void main(String[] args) {
		launch(args);
	}
}
