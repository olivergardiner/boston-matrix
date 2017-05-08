package uk.org.whitecottage.bostonmatrix;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generate {

	private static final Logger LOGGER = Logger.getLogger(Generate.class.getName());

	private Generate() {
		
	}

	public static void main(String[] args) {
		String file = "Boston_Matrix";
		
		if (args.length > 0) {
			file = args[0];
		}
		
		FileInputStream input = null;
		XSSFWorkbook xlsx = null;
		
		try {
			input = new FileInputStream(new File("xlsx/" + file + ".xlsx"));
			xlsx = new XSSFWorkbook(input);
		} catch (Exception e) {
			LOGGER.log(Level.SEVERE, e.getMessage(), e);
		} finally {
			try {
				if (input != null) {
					input.close();
				}
			} catch (IOException e) {
				LOGGER.log(Level.SEVERE, e.getMessage(), e);
			}
		}
		
		BostonMatrix bostonMatrix = new BostonMatrix(xlsx);
		
		FileOutputStream output = null;
		
		try {
			output = new FileOutputStream(new File("pptx/" + file + ".pptx"));
			bostonMatrix.writeSlides(output);
		} catch (Exception e) {
			LOGGER.log(Level.SEVERE, e.getMessage(), e);
		} finally {
			try {
				if (output != null) {
					output.close();
				}
			} catch (IOException e) {
				LOGGER.log(Level.SEVERE, e.getMessage(), e);
			}
		}
	}
}
