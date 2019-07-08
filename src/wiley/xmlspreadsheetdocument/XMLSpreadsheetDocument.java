package wiley.xmlspreadsheetdocument;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.xml.stream.XMLStreamException;

import wiley.xmlspreadsheetlibrary.Cell;
import wiley.xmlspreadsheetlibrary.CellStyle;
import wiley.xmlspreadsheetlibrary.Font;
import wiley.xmlspreadsheetlibrary.Row;
import wiley.xmlspreadsheetlibrary.Sheet;
import wiley.xmlspreadsheetlibrary.Workbook;

public class XMLSpreadsheetDocument {
	private Workbook workbook;
	private Sheet selectedSheet;
	
	public XMLSpreadsheetDocument(String filePath) throws FileNotFoundException, XMLStreamException {
		workbook = new Workbook(filePath);
	}
	
	public XMLSpreadsheetDocument() {
		workbook = new Workbook();
	}
	
	public void createSheet(String sheetName) {
		workbook.createSheet(sheetName);
	}
	
	public void selectSheetByName(String sheetName) {
		selectedSheet = workbook.getSheetByName(sheetName);
	}
	
	public void selectSheetByIndex(int sheetIndex) {
		selectedSheet = workbook.getSheetByIndex(sheetIndex);
	}
	
	public String getValue(int rowIndex, int cellIndex) {
		Row row = selectedSheet.getRow(rowIndex);
		if(row == null) {
			return null;
		}
		Cell cell = row.getCell(cellIndex);
		if(cell == null) {
			return null;
		}
		return cell.getValue();
	}
	
	public String getValue(int rowIndex, int cellIndex, String defaultValue) {
		Row row = selectedSheet.getRow(rowIndex);
		if(row == null) {
			return defaultValue;
		}
		Cell cell = row.getCell(cellIndex);
		if(cell == null) {
			return defaultValue;
		}
		if(cell.getValue() == null) {
			return defaultValue;
		}
		return cell.getValue();
	}
	
	public void setValue(int rowIndex, int cellIndex, String value) { // произветси создание строчек и столбцов по индексу
		Row row = selectedSheet.getRow(rowIndex);
		
		if(row == null) {
			row = selectedSheet.createRow(); // произветси создание строчек и столбцов по индексу
		}
		
		Cell cell = row.getCell(cellIndex);
		
		if(cell == null) {
			cell = row.createCell(); // произветси создание строчек и столбцов по индексу
		}
		
		cell.setValue(value);
	}
	
	public void setValue(int rowIndex, int cellIndex, String value, ExcelStyle style) { // произветси создание строчек и столбцов по индексу
		Row row = selectedSheet.getRow(rowIndex);
		
		if(row == null) {
			row = selectedSheet.createRow(); // произветси создание строчек и столбцов по индексу
		}
		
		Cell cell = row.getCell(cellIndex);
		
		if(cell == null) {
			cell = row.createCell(); // произветси создание строчек и столбцов по индексу
		}
		
		cell.setStyle(style.getCellStyle());
		cell.setValue(value);
	}
	
	public void setValues(int rowIndex, String[] values) {
		for (int i = 0; i < values.length; i++) {
			setValue(rowIndex, i, values[i]);
		}
	}
	
	public void setValues(int rowIndex, String[] values, ExcelStyle style) {
		for (int i = 0; i < values.length; i++) {
			setValue(rowIndex, i, values[i], style);
		}
	}
	
	public void addValues(String[] values) {
		selectedSheet.createRow();
		for (int i = 0; i < values.length; i++) {
			setValue(selectedSheet.getRowsCount()-1, i, values[i]);
		}
	}
	
	public void addValues(String[] values, ExcelStyle style) {
		selectedSheet.createRow();
		for (int i = 0; i < values.length; i++) {
			setValue(selectedSheet.getRowsCount()-1, i, values[i], style);
		}
	}
	
	public int getRowCount() {
		return selectedSheet.getRowsCount();
	}
	
	public void write(String filePath) throws IOException {
		workbook.write(filePath);
	}
	
	public void setColumnWidth(int columnIndex, double width) {
		selectedSheet.setColumnWidth(columnIndex, width);
	}
	
	public ExcelStyle createStyle() {
		CellStyle c = workbook.createStyle();
		return new ExcelStyle(c);
	}
	
	public ExcelFont createFont() {
		Font f = workbook.createFont();
		return new ExcelFont(f);
	}
	
	/*private CellStyle convertStyle(ExcelStyle st, CellStyle c) {
		if(st.getFont() != null) {
			Font f = workbook.createFont();
			if(st.getFont().isBold()) {
				f.setBold(true);
			}
			c.setFont(f);
		}
		
		if(st.getBackgroundColor() != null) {
			c.setBackgroundColor(String.format("#%02x%02x%02x", st.getBackgroundColor().getRed(), st.getBackgroundColor().getGreen(), st.getBackgroundColor().getBlue()));
		}
		return c;
	}*/
}
