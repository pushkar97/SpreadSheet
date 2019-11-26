package spreadsheets;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This is a convinience class for creating and reading microsoft excel spreadsheets.
 * It takes care of most of the edge cases when working with poi and tries to convert string data to it's most appropriate datatype
 * It also supports both .xls and .xlsx formats. and supports limited styling with html like syntax
 * <pre>
 * SpreadSheet s = new SpreadSheet("data.xlsx");
 *
 *		List<List<Object>> data = new ArrayList<List<Object>>();
 *		List<Object> row1 = new ArrayList<Object>();
 *		row1.add("h1<html/>");
 *		row1.add("245<POI-UNDERLINE/>");
 *		row1.add("256D<POI-BOLD/>");
 *		
 *		s.writeData("Sheet1", 0, 0, data);
 *		
 *		s.saveWorkbook();
 *		s.closeWorkbook();
 * </pre>
 * @author Pushkar
 */
public class SpreadSheet {
	private Workbook workbook;
	private File file;
	private String fileExtensionName;
	private int rowid;
	private int cellid;
	private CreationHelper createHelper;
	private CellStyle hyperlinkStyle;
	private Font hyperlinkFont;
	public int getRowid() {
		return rowid;
	}

	public int getCellid() {
		return cellid;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public File getFile() {
		return file;
	}

	public SpreadSheet(String filepath) throws IOException {
		this(new File(filepath));
	}

	public SpreadSheet(File file) throws IOException {
		if (!file.exists()) {
			file.createNewFile();
		}
		this.file = file;
		createWorkbook();
		setDefaults();
	}
	
	private void createWorkbook() throws IOException {
		this.fileExtensionName = this.file.getName().substring(this.file.getName().indexOf("."));
		if (file.length() == 0) {
			switch (this.fileExtensionName) {
			case ".xlsx":
				workbook = new XSSFWorkbook();
				break;
			case ".xls":
				workbook = new HSSFWorkbook();
				break;
			default:
				throw new IOException("Invalid file type");
			}
		} else {
			FileInputStream inputStream = new FileInputStream(file);

			switch (this.fileExtensionName) {
			case ".xlsx":
				workbook = new XSSFWorkbook(inputStream);
				break;
			case ".xls":
				workbook = new HSSFWorkbook(inputStream);
				break;
			default:
				throw new IOException("Invalid file type");
			}
		}
		System.out.println("File successfully loaded : "+file.getAbsolutePath());
	}
	
	private void setDefaults() {
		createHelper = workbook.getCreationHelper();
		hyperlinkStyle = workbook.createCellStyle();
		hyperlinkFont = workbook.createFont();
		hyperlinkFont.setUnderline(Font.U_SINGLE);
		hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
		hyperlinkStyle.setFont(hyperlinkFont);
	}
	
	/**
	 * @param sheetname - Sheet name to which data will be written. if sheet is present data will be overwritten, if not new sheet will be created
	 * @param initRowId - row number from which data will be written, useful while making multiple write calls to same sheet.
	 * @param initCellId - initial cell id, useful while making multiple write calls to same sheet.
	 * @param data - Two dimensional data which will be written to sheet.
	 */
	public void writeData(String sheetname, int initRowId, int initCellId, List<List<Object>> data) {
		Sheet spreadsheet = workbook.getSheet(sheetname);
		if (spreadsheet == null)
			spreadsheet = workbook.createSheet(sheetname);
		
		rowid = initRowId;
		for(List<Object> rowSet: data) {
			Row row = spreadsheet.createRow(rowid++);
			if(rowSet == null || rowSet.size() == 0)
				continue;
			cellid = initCellId;
			for(Object obj : rowSet) {
				Cell cell = row.createCell(cellid++);
				setCellValue(cell,obj);
			}
		}
		System.out.println("Writing completed successfully");
	}
	
	private String compileCell(Cell cell, String cellValue) {
		String tagRegex = "\\<[Pp][Oo][Ii]-.+?/\\>";
		Pattern tagPattern = Pattern.compile(tagRegex);
		Matcher tagMatcher = tagPattern.matcher(cellValue);
		List<String> tagMatches = new ArrayList<String>();
		while (tagMatcher.find()) {
			tagMatches.add(tagMatcher.group());
		}
		if(tagMatches.isEmpty()) {
			return cellValue;
		}
		CellStyle style;
		Font font;
		style = workbook.createCellStyle();
		font = workbook.createFont();
		for (String tag : tagMatches) {
			if (tag.toUpperCase().matches("\\<POI-HYPERLINK.*\\>")) {
				Map<String,String> attrs = getAttributes(tag.toUpperCase());
				System.out.println(attrs.keySet().toString());
				if(attrs.containsKey("TYPE") && attrs.containsKey("URL")) {
					try {
						Hyperlink link = createHelper.createHyperlink(HyperlinkType.valueOf(attrs.get("TYPE")));
			            link.setAddress(attrs.get("URL"));
			            cell.setCellStyle(hyperlinkStyle);
			            cell.setHyperlink(link);
			    		return cellValue.replaceAll(tagRegex, "");
					}catch(IllegalArgumentException e) {}
				}
			}
			if (tag.toUpperCase().equals("<POI-BOLD/>")) {
				font.setBold(true);
			}
			if (tag.toUpperCase().equals("<POI-ITALIC/>")) {
				font.setItalic(true);
			}
			if (tag.toUpperCase().equals("<POI-STRIKEOUT/>")) {
				font.setStrikeout(true);
			}
			if (tag.toUpperCase().equals("<POI-UNDERLINE/>")) {
				font.setUnderline(Font.U_SINGLE);
			}
			if (tag.toUpperCase().matches("\\<POI-FONTCOLOR\\s*=\\s*\"\\w+\".*/\\>")) {
				try {
					font.setColor(IndexedColors.valueOf(tag.substring(tag.indexOf('"')+1, tag.lastIndexOf('"')).toUpperCase()).index);
				}catch(IllegalArgumentException e) {
				}
			}
			style.setFont(font);
			if (tag.toUpperCase().matches("\\<POI-BGCOLOR\\s*=\\s*\"\\w+\".*/\\>")) {
				try {
					style.setFillForegroundColor(IndexedColors.valueOf(tag.substring(tag.indexOf('"')+1, tag.lastIndexOf('"')).toUpperCase()).index);
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}catch(IllegalArgumentException e) {
				}
			}
			if (tag.toUpperCase().matches("\\<POI-BORDER.*\\>")) {
				Map<String,String> attrs = getAttributes(tag.toUpperCase());
				if(attrs.containsKey("ALL")) {
					try {
						style.setBorderTop(BorderStyle.valueOf(attrs.get("ALL")));
						style.setBorderRight(BorderStyle.valueOf(attrs.get("ALL")));
						style.setBorderBottom(BorderStyle.valueOf(attrs.get("ALL")));
						style.setBorderLeft(BorderStyle.valueOf(attrs.get("ALL")));
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("TOP")) {
					try {
						style.setBorderTop(BorderStyle.valueOf(attrs.get("TOP")));
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("LEFT")) {
					try {
						style.setBorderLeft(BorderStyle.valueOf(attrs.get("LEFT")));
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("BOTTOM")) {
					try {
						style.setBorderBottom(BorderStyle.valueOf(attrs.get("BOTTOM")));
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("RIGHT")) {
					try {
						style.setBorderRight(BorderStyle.valueOf(attrs.get("RIGHT")));
					}catch(IllegalArgumentException e) {}
				}
				
				if(attrs.containsKey("COLOR")) {
					try {
						style.setTopBorderColor(IndexedColors.valueOf(attrs.get("COLOR")).index);
						style.setRightBorderColor(IndexedColors.valueOf(attrs.get("COLOR")).index);
						style.setBottomBorderColor(IndexedColors.valueOf(attrs.get("COLOR")).index);
						style.setLeftBorderColor(IndexedColors.valueOf(attrs.get("COLOR")).index);
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("COLOR_TOP")) {
					try {
						style.setTopBorderColor(IndexedColors.valueOf(attrs.get("COLOR_TOP")).index);
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("COLOR_RIGHT")) {
					try {
						style.setRightBorderColor(IndexedColors.valueOf(attrs.get("COLOR_RIGHT")).index);
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("COLOR_BOTTOM")) {
					try {
						style.setBottomBorderColor(IndexedColors.valueOf(attrs.get("COLOR_BOTTOM")).index);
					}catch(IllegalArgumentException e) {}
				}else if(attrs.containsKey("COLOR_LEFT")) {
					try {
						style.setLeftBorderColor(IndexedColors.valueOf(attrs.get("COLOR_LEFT")).index);
					}catch(IllegalArgumentException e) {}
				}
			}
		}
		cell.setCellStyle(style);
		return cellValue.replaceAll(tagRegex, "");
	}
	
	private Map<String,String> getAttributes(String tag){
		String attrRegex = "\\w+\\s*=\\s*\".+?\"";
		Pattern attrPattern = Pattern.compile(attrRegex);
		Matcher tagMatcher = attrPattern.matcher(tag);
		Map<String,String> attrMatches = new HashMap<String,String>();
		while (tagMatcher.find()) {
			String attr = tagMatcher.group();
			String attrName = attr.substring(0, attr.indexOf('=')).trim();
			String attrVal = attr.substring(attr.indexOf('"')+1, attr.indexOf('"',attr.indexOf('"')+1)).trim();
			attrMatches.put(attrName, attrVal);
		}
		return attrMatches;
	}
	
	private Cell setCellValue(Cell cell,Object obj) {
		if(obj instanceof Date) {
			cell.setCellValue(((Date)obj));
		}else if(obj instanceof Double || obj instanceof Float) {
			cell.setCellValue(((Number)obj).doubleValue());
		}else if(obj instanceof Number) {
			cell.setCellValue(((Number)obj).longValue());
		}else if (obj instanceof Boolean) {
			cell.setCellValue((Boolean.valueOf(String.valueOf(obj).trim())).booleanValue());
		}else if (obj instanceof String) {
			String value = compileCell(cell,(String)obj);
			try {
				Date date1= new SimpleDateFormat("yyyy-MM-dd").parse(value);
				cell.setCellValue(date1);
			}catch(ParseException e2) {
				try {
					cell.setCellValue(Long.valueOf(value));
				}catch(NumberFormatException e) {
					try{
						cell.setCellValue(Double.valueOf(value));
					}catch(NumberFormatException e1) {
						if(value.trim().equalsIgnoreCase("true") || value.trim().equalsIgnoreCase("false")) {
							cell.setCellValue((Boolean.valueOf(value.trim())).booleanValue());
						}else {
							cell.setCellValue(value);
						}
					}
				}
			}
		}else {
			cell.setCellValue(String.valueOf(obj));
		}
		return cell;
	}

	/**
	 * @param sheetName - sheet from which data will be read.
	 * @return - returns 2 dimensional data as a List<List<Object>>
	 * @throws IOException - If sheet with provided sheetname is not present
	 */
	public List<List<Object>> getData(String sheetName) throws IOException {
		System.out.println("Reading data..." + sheetName);
		Sheet sheet = workbook.getSheet(sheetName);
		if (sheet == null) {
			throw new IOException("Invalid Sheet Name : ".concat(sheetName));
		}
		List<List<Object>> sheetArray = new ArrayList<List<Object>>();
		int colLen = sheet.getRow(0).getLastCellNum();
		sheet.forEach(row -> {
			if (row != null) {
				List<Object> rowArray = new ArrayList<Object>();
				int lastColumn = Math.max(colLen, row.getLastCellNum());
				for (int i = 0; i < lastColumn; i++) {
					Cell cell = row.getCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK);
					rowArray.add(getCellValue(cell));
				}
				sheetArray.add(rowArray);
			}
		});
		return sheetArray;
	}
	
	private Object getCellValue(Cell cell) {
		Object obj = "";
		switch (cell.getCellType()) {
		case BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case STRING:
			obj = cell.getRichStringCellValue().getString();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				obj = cell.getDateCellValue();
			} else {
				obj = cell.getNumericCellValue();
			}
			break;
		case FORMULA:
			obj = cell.getCellFormula();
			break;
		case BLANK:
			obj = "";
			break;
		case ERROR:
			obj = cell.getErrorCellValue();
			break;
		default:
			obj = "";
			break;
		}
		return obj;
	}

	public void saveWorkbook() throws IOException {
		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("File saved successfully");
	}
	
	public void closeWorkbook() throws IOException {
		try {
			workbook.close();
		} catch (IOException e) {
			System.out.println("Error closing workbook");
			throw new IOException();
		}
	}

	/**
	 * Main method with data write test showing behaviour of the code
	 * @param args
	 * @throws IOException
	 * @throws ParseException
	 */
	public static void main(String... args) throws IOException, ParseException {
		SpreadSheet sheet = new SpreadSheet("D:\\Book1.xlsx");
		List<List<Object>> data = new ArrayList<List<Object>>();
		List<Object> row1 = new ArrayList<Object>();
		row1.add("h1<html/>");
		row1.add("245<POI-UNDERLINE/>");
		row1.add("256D<POI-BOLD/>");
		List<Object> row2 = new ArrayList<Object>();
		row2.add(null);
		row2.add("d2<poi-italic/>");
		row2.add("h3<POI-BGCOLOR = \"Gold\"/>");
		List<Object> row3 = new ArrayList<Object>();
		List<Object> row4 = new ArrayList<Object>();
		row4.add("true<POI-BORDER ALL=\"MEDIUM_DASHED\" COLOR=\"TEAL\"/>");
		row4.add("256<POI-FONTCOLOR=\"Green\"/><POI-BGCOLOR=\"LIGHT_Green\"/>");
		row4.add("1997-11-25<POI-BOLD/>");
		row4.add(new SimpleDateFormat("yyyy-MM-dd").parse("1997-05-17"));
		
		List<Object> row5 = new ArrayList<Object>();
		row5.add("true<POI-BORDER ALL=\"MEDIUM_DASHED\" COLOR=\"TEAL\"/>");
		row5.add("pushkar97<POI-HYPERLINK TYPE=\"URL\" URL=\"https://github.com/pushkar97\" />");
		row5.add("1997-11-25<POI-BOLD/>");
		row5.add(new SimpleDateFormat("yyyy-MM-dd").parse("1997-05-17"));
		data.add(row1);
		data.add(null);
		data.add(row2);
		data.add(row3);
		data.add(row4);
		data.add(row4);
		data.add(row5);
		sheet.writeData("dataWriteTest", 0, 0, data);
		sheet.saveWorkbook();
		sheet.closeWorkbook();
		data = sheet.getData("dataWriteTest");
		data.forEach(row -> {
			row.forEach(col -> {
				System.out.print(col+"\t");
			});
			System.out.println(row.size());
		});
	}
}
