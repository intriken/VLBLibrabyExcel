package dox.vlb.uams;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.FieldPosition;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import vlb.ide.utils.VLBDateUtils;

public class POIUtils {
	
	public static FormulaEvaluator formulaEvaluation(Cell cell) {
		Workbook wbook = cell.getSheet().getWorkbook();
	    return wbook.getCreationHelper().createFormulaEvaluator();
	}
	
	public static XSSFFont boldfont(XSSFWorkbook wb){
		XSSFFont bold = null;
		if(bold == null){
			bold = wb.createFont();
			bold.setFontHeightInPoints((short)11);
			bold.setBoldweight(Font.BOLDWEIGHT_BOLD);		
		}
		return bold;
	}

	public static void createNamedRange(Workbook wb, String sheetName, String namename, int srow, int scol, int erow, int ecol) {
		Name namedCel = wb.createName();
		namedCel.setNameName(namename);
		erow++;
		srow++;
		String reference = "'" + sheetName +"'!"+ CellReference.convertNumToColString(scol) + srow + ":" + CellReference.convertNumToColString(ecol) + erow; 		    
		namedCel.setRefersToFormula(reference);

	}

	public static Cell createCell(Sheet sheet, int rown, int coln) {
		Row row = sheet.getRow(rown);
		if(row == null){
			row = sheet.createRow(rown);
		}
		Cell cell = row.getCell(coln);
		if(cell == null){
			cell = row.createCell(coln);
		}
		return cell;
	}
	
	public static String getStringValue(Sheet sheet, int row, int col){	
		try{
			return sheet.getRow(row).getCell(col).toString();
		} catch (NullPointerException e){
			return "";
		}

	}
	
	public static String getStringValue(Row row, int col){	
		Cell cell = row.getCell(col);		
		return getStringValue(cell);
	}
	
	public static void setCellComment(Cell cell, String message, int size) {
	    Drawing drawing = cell.getSheet().createDrawingPatriarch();
	    CreationHelper factory = cell.getSheet().getWorkbook()
	            .getCreationHelper();
	    // When the comment box is visible, have it show in a 1x3 space
	    ClientAnchor anchor = factory.createClientAnchor();
	    anchor.setCol1(cell.getColumnIndex());
	    anchor.setCol2(cell.getColumnIndex() + 5);
	    anchor.setRow1(cell.getRowIndex());
	    anchor.setRow2(cell.getRowIndex() + size);	    
	    anchor.setDx1(100 * XSSFShape.EMU_PER_PIXEL);
	    anchor.setDx2(100 * XSSFShape.EMU_PER_PIXEL);
	    anchor.setDy1(100 * XSSFShape.EMU_PER_PIXEL);
	    anchor.setDy2(100 * XSSFShape.EMU_PER_PIXEL);

	    // Create the comment and set the text+author
	    Comment comment = drawing.createCellComment(anchor);
	    RichTextString str = factory.createRichTextString(message);
	    comment.setString(str);
	    comment.setAuthor("Apache POI");
	    // Assign the comment to the cell
	    cell.setCellComment(comment);
	}
	
	public static String getStringValue(Cell cell){

		if(cell == null){
			return "";
		}

		DateFormat df = new SimpleDateFormat("MM/dd/yyyy");

		if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			if(DateUtil.isCellDateFormatted(cell)) {
				Date date = cell.getDateCellValue();
				return VLBDateUtils.dateformat("dd-MMM-yyyy", date);
			} else {
				double value = cell.getNumericCellValue();
				return getDoubleFormated(value);
			}
		} else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			CellValue cellValue = formulaEvaluation(cell).evaluate(cell);	

			String svalue = cellValue.getStringValue();
			if(svalue != null){
				if(!svalue.isEmpty()){
					return svalue;
				}
			}

			double dvalue = cellValue.getNumberValue();
			if(cellValue.getCellType() == Cell.CELL_TYPE_NUMERIC){
				if (isCellDateFormatted(cell, cellValue)){
					return VLBDateUtils.dateformat("dd-MMM-yyyy", DateUtil.getJavaDate(dvalue));				
				} else {					
					return getDoubleFormated(dvalue);					
				}		
			}

			return "";		
		} else {
			try {
				Date da = df.parse(cell.toString().trim());
				return VLBDateUtils.dateformat("dd-MMM-yyyy", da);
			} catch (ParseException e) {}			

			return cell.toString();
		}

	}
	


	public static boolean isCellDateFormatted(Cell cell, CellValue cellvalue) {
		if (cellvalue == null) {
			return false;
		}
		boolean bDate = false;

		double d = cellvalue.getNumberValue();
		if ( DateUtil.isValidExcelDate(d) ) {
			CellStyle style = cell.getCellStyle();
			if(style==null) return false;
			int i = style.getDataFormat();
			String f = style.getDataFormatString();
			bDate = DateUtil.isADateFormat(i, f);
		}
		return bDate;
	}

	public static String getDoubleFormated(double value){
		String formattingString = "#0.#####";
		DecimalFormat formatter = new DecimalFormat(formattingString);
		FieldPosition fPosition = new FieldPosition(0);
		StringBuffer buffer = new StringBuffer();
		return formatter.format(value, buffer, fPosition).toString();
	}

	public static Workbook createWbr(String fileloc) {
		Workbook wbr = null;
		InputStream inp = null;
		try {
			inp = new FileInputStream(fileloc);
		} catch (FileNotFoundException e1) {
			System.out.println("No File Exists");
		}

		try {
			wbr = WorkbookFactory.create(inp);
		} catch (InvalidFormatException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		return wbr;
	}
}
