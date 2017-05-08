package uk.org.whitecottage.bostonmatrix;

import static uk.org.whitecottage.poi.ShapeHelper.circle;
import static uk.org.whitecottage.poi.ShapeHelper.line;
import static uk.org.whitecottage.poi.ShapeHelper.simpleText;
import static uk.org.whitecottage.poi.SheetHelper.getCellStringValue;
import static uk.org.whitecottage.poi.SheetHelper.nextRow;

import java.awt.Dimension;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BostonMatrix {
	protected XSSFWorkbook xlsx;
	protected double scale = 0.7;
	
	protected double x0;
	protected double y0;
	protected double size;
	
	protected static final double MAX_RADIUS = 50;
	protected static final double MIN_RADIUS = 5;

	public BostonMatrix(XSSFWorkbook xlsx) {
		this.xlsx = xlsx;
	}
	
	public void writeSlides(OutputStream output) throws IOException {
		XMLSlideShow pptx = new XMLSlideShow();
		
		buildSlide(pptx.createSlide());
		
		try {
			pptx.write(output);
		} catch (IOException e) {
			throw e;
		} finally {
			pptx.close();
		}
	}
	
	protected void buildSlide(XSLFSlide slide) {
		Dimension pageSize = slide.getSlideShow().getPageSize();
		int pageWidth = pageSize.width;
		int pageHeight = pageSize.height;
		
		Sheet matrix = xlsx.getSheet("Boston Matrix");
		
		setScale(pageHeight, pageWidth);
		
		addFurniture(matrix, slide);
		
		List<Entry> entries = new ArrayList<>();
		double max = 0;
		
		Row row = matrix.getRow(10);
		for (String name = getCellStringValue(row, 2); !"".equals(name); name = getCellStringValue(row = nextRow(row), 2)) {
			Entry entry = new Entry(row);
			entries.add(entry);
			double z = entry.getZ();
			if (!Double.isNaN(z) && z > max) {
				max = z;
			}
		}
		
		entries.sort(null);
		
		for (Entry entry: entries) {
			plotEntry(slide, entry, max);
		}
		
		/*XSLFTextBox box = slide.createTextBox();
		box.setText("Box");
		box.setAnchor(new Rectangle(pageWidth/2, pageHeight/2, pageWidth/10, pageWidth/10));
		box.setFillColor(Color.blue);
		box.setLineColor(Color.black);;
		
		XSLFAutoShape shape = slide.createAutoShape();
		shape.setShapeType(ShapeType.ELLIPSE);
		shape.setAnchor(new Rectangle(pageWidth/4, pageHeight/4, pageWidth/5, pageWidth/10));
		shape.setFillColor(Color.blue);
		shape.setStrokeStyle(1.0);*/
	}
	
	protected void setScale(int pageHeight, int pageWidth) {
		size = pageHeight * scale; 
		x0 = (pageWidth - size) / 2;
		y0 = (pageHeight - size) / 2;
	}
	
	protected void addFurniture(Sheet matrix, XSLFSlide slide) {
		
		String title = matrix.getRow(1).getCell(2).getStringCellValue();
		String xAxis = matrix.getRow(2).getCell(2).getStringCellValue();
		String yAxis = matrix.getRow(3).getCell(2).getStringCellValue();
		//String quadrant11 = matrix.getRow(4).getCell(2).getStringCellValue();
		//String quadrant12 = matrix.getRow(5).getCell(2).getStringCellValue();
		//String quadrant21 = matrix.getRow(6).getCell(2).getStringCellValue();
		//String quadrant22 = matrix.getRow(7).getCell(2).getStringCellValue();
		
		
		line(slide, x0, y0, x0, y0 + size);
		line(slide, x0, y0 + size, x0 + size, y0 + size);
		line(slide, x0 + size, y0 + size, x0 + size, y0);
		line(slide, x0 + size, y0, x0, y0);

		line(slide, x0 + size / 2, y0, x0 + size / 2, y0 + size);
		line(slide, x0, y0 + size / 2, x0 + size, y0 + size / 2);
		
		simpleText(slide, x0 + size / 2, y0 - 20, title, 20.0);
		simpleText(slide, x0 + size / 2, y0 + size + 12, xAxis, 12.0);
		simpleText(slide, x0 - 12, y0 + size / 2, yAxis, 12.0, 270.0, TextParagraph.TextAlign.CENTER);
	}
	
	protected void plotEntry(XSLFSlide slide, Entry entry, double max) {
		double z = entry.getZ();
		
		if (z < 0 || Double.isNaN(z)) {
			z = 0.0;
		}
		
		if (max <= 0) {
			z = MIN_RADIUS;
		} else {
			z = (z / max) * (MAX_RADIUS - MIN_RADIUS) + MIN_RADIUS;
		}
		
		double x = entry.getX() * size / 100 + x0;
		double y = (100 - entry.getY()) * size / 100 + y0;
		
		circle(slide, x, y, z);
		simpleText(slide, x + z / 2 + 5, y, entry.getName(), 12.0, 0.0, TextParagraph.TextAlign.LEFT);
	}
}
