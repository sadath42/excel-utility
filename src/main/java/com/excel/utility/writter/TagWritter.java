package com.excel.utility.writter;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.excel.utility.exception.UtilityException;

public class TagWritter {
	private static Logger LOGGER = LoggerFactory.getLogger(MethodHandles.lookup().lookupClass().getSimpleName());

	public void writetoExcel(String excelPath, String inputFile) {

		LOGGER.info("Writing to the file {} from file {} ", excelPath, inputFile);
		// Creating a Workbook from an Excel file (.xls or .xlsx)
		DataFormatter dataFormatter = new DataFormatter();

		// BufferedReader object for input.txt
		try (BufferedReader br = new BufferedReader(new FileReader(inputFile))) {
			FileInputStream fileInputStream = new FileInputStream(excelPath);
			Workbook workbook = WorkbookFactory.create(fileInputStream);

			Sheet sheet = workbook.getSheetAt(0);

			int count = 1;
			String line = br.readLine();
			// loop for each line of input.txt
			while (line != null) {

				Row row = sheet.getRow(count);
				if (row == null || row.getCell(0) == null || dataFormatter.formatCellValue(row.getCell(0)).isEmpty()) {
					break;
					// row = sheet.createRow(count);
				}
				int i = 1;
				while (i <= 5 && line != null) {
					line = line.trim();
					if (line.isEmpty()) {
						line = br.readLine();
						continue;
					}
					int tagCount = 3 + i;
					// Update the value of cell
					Cell cell = row.getCell(tagCount);
					if (cell == null) {
						cell = row.createCell(tagCount);
					}
					cell.setCellValue(line);
					line = br.readLine();

					i++;
				}
				count++;

			}

			fileInputStream.close();
			FileOutputStream outFile = new FileOutputStream(new File(excelPath));
			workbook.write(outFile);
			outFile.close();
			workbook.close();

		} catch (EncryptedDocumentException | IOException e) {
			throw new UtilityException(e.getMessage(), e);
		}

	}

	public void removeDuplicateAndShuffle(String filePath, boolean isShuffle, boolean dontRemoveDuplicates) {
		LOGGER.info("isShuffle {}  dontRemoveDuplicates {} ", isShuffle, dontRemoveDuplicates);
		// PrintWriter object for output.txt
		try (BufferedWriter writer = new BufferedWriter(new FileWriter("output.txt"));
				BufferedReader br = new BufferedReader(new FileReader(filePath))) {
			String line = br.readLine();

			// set store unique values
			HashSet<String> hs = new HashSet<String>();
			List<String> fileContents = new ArrayList<>();

			// loop for each line of input.txt
			while (line != null) {
				// write only if not
				// present in hashset

				if (dontRemoveDuplicates || hs.add(line.toLowerCase())) {
					line = line.trim();
					if (!line.isEmpty()) {
						fileContents.add(line);
					}
				}

				line = br.readLine();

			}

			if (isShuffle) {
				Collections.shuffle(fileContents);
			}
			for (String unique : fileContents) {
				writer.write(unique);
				writer.newLine();
			}
			writer.flush();
		} catch (IOException e) {
			throw new UtilityException(e.getMessage(), e);
		}

	}

}
