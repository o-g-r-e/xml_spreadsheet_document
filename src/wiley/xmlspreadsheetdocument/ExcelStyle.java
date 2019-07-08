package wiley.xmlspreadsheetdocument;
import java.awt.Color;

import wiley.xmlspreadsheetlibrary.CellStyle;

public class ExcelStyle {
	
	private CellStyle cellStyle;
	
	protected ExcelStyle() {}
	
	protected ExcelStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	
	public void setBackgroundColor(Color backgroundColor) {
		cellStyle.setBackgroundColor(String.format("#%02x%02x%02x", backgroundColor.getRed(), backgroundColor.getGreen(), backgroundColor.getBlue()));
	}
	
	public void setFont(ExcelFont font) {
		cellStyle.setFont(font.getFont());
	}
	
	public String getId() {
		return cellStyle.getId();
	}
	
	protected CellStyle getCellStyle() {
		return cellStyle;
	}
	
	/*private String id;
	private Color backgroundColor;
	private Color borderColor;
	private ExcelFont font;
	
	public Color getBackgroundColor() {
		return backgroundColor;
	}
	public void setBackgroundColor(Color backgroundColor) {
		this.backgroundColor = backgroundColor;
	}
	public Color getBorderColor() {
		return borderColor;
	}
	public void setBorderColor(Color borderColor) {
		this.borderColor = borderColor;
	}
	public ExcelFont getFont() {
		return font;
	}
	public void setFont(ExcelFont font) {
		this.font = font;
	}
	public String getId() {
		return id;
	}
	
	
	
	*/
}
