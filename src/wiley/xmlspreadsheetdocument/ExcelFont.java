package wiley.xmlspreadsheetdocument;

import wiley.xmlspreadsheetlibrary.Font;

public class ExcelFont {
	
	private Font font;
	
	protected ExcelFont() {}
	
	protected ExcelFont(Font font) {
		this.font = font;
	}
	
	public void setBold(boolean bold) {
		font.setBold(bold);
	}
	
	public void setItalic(boolean italic) {
		font.setItalic(italic);
	}

	public Font getFont() {
		return font;
	}
	
	/*private boolean bold;
	private boolean italic;
	
	public boolean isBold() {
		return bold;
	}
	public void setBold(boolean bold) {
		this.bold = bold;
	}
	public boolean isItalic() {
		return italic;
	}
	public void setItalic(boolean italic) {
		this.italic = italic;
	}*/
}
