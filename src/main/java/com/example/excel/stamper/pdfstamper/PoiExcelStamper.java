package com.example.excel.stamper.pdfstamper;

import java.util.ArrayList;
import java.util.List;
import java.util.Arrays;

import org.springframework.stereotype.Component;

import java.nio.channels.ClosedSelectorException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


@Component
public class PoiExcelStamper  {
	String fileLocation = "Test Domestic Standard Upload Template.xlsx";
	Workbook workbook = null;

	public PoiExcelStamper() {
		
		try {
			FileInputStream inputStream = new FileInputStream(new File(fileLocation));
			workbook = new XSSFWorkbook(inputStream);
			FormulaEvaluator evaluator = workbook.getCreationHelper()
				.createFormulaEvaluator();

		} catch (Exception e) {
			e.printStackTrace();	
		}
	}

	// @method getJsonFromNamedCols 
	// reads list of named fields from a worksheet (by row)
	// it assumes all subsequent cols in names are in the same worksheet as first  
	// <p>
	// @param names a list of named columns to read
	// @return a json doc

	public void getJsonFromNamedCols(List<String> names) {

		if (names.size() == 0) return;

		CellReference cellReference = getNamedCell(names.get(0));
		Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
		int startRow = cellReference.getRow();
		Row r = null; 

		do {
			r =  workSheet.getRow(++startRow); 

			if (r != null) {

				for(String val : names) {
				
					CellReference cellsref = getNamedCell(val);				
					Cell c = r.getCell(cellsref.getCol()); 
					
					if (c != null)
					switch (c.getCellType()) {
						case NUMERIC:
							System.out.print(">>> " + val + ":" + c.getNumericCellValue());
							continue;
						case STRING:
							System.out.print(">>> " + val + ":" +c.getStringCellValue());
							continue;
						case BLANK:
						case _NONE:
						case ERROR:
							break;
					} 			
				}		
				System.out.print("\n");
			} 
		} while (r != null);
	} 

	// @method getJsonFromNamedCols 
	// reads list of named fields from a worksheet (by row)
	// it assumes all subsequent cols in names are in the same worksheet as first  
	// <p>
	// @param names a list of named columns to read
	// @return a json doc

	public void writeToNamedCols(List<String> names) 
		throws Exception {

		try {

			CellReference cellReference = getNamedCell(names.get(0));
			Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
			Row headerRow = workSheet.getRow(cellReference.getRow());
			//int startRow = cellReference.getRow();
			Row r = null; 
			int startRow = workSheet.getLastRowNum(); 
			System.out.println(">>> write-cols " + names);

			do {
				r = workSheet.createRow(++startRow); // workSheet.getRow(++startRow); 

				if (r != null) {

					for(String val : names) {		
						Name namedCell = workbook.getName(val); 
						CellReference cellsref = new CellReference(namedCell.getRefersToFormula());
						Cell c = r.createCell(cellsref.getCol());
						CellStyle cs = headerRow.getCell(cellsref.getCol()).getCellStyle();
						cs.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);  
						c.setCellStyle(cs);
						c.setCellValue((String) "val:" + val);					
					}
				}
			} while (startRow < 20);

			writeXlsxFile("modified-3.xlsx");

		} catch (Exception e) {
			throw e;
		}
	} 

	// @method getNames 
	// @param names a list of names to check within the workbook
	// @return a list of names present within the workbook 
	//
	public List<String> getNames(List<String> names) {
		List<String> present = new ArrayList<String>();		
		//java.util.List<? extends Name> allNames = workbook.getAllNames();
		
		for(String val : names){
			if (workbook.getName(val) != null)
				present.add(val); 
		}
		return present;
	} 

	// @method getNamedCell 
	// @return a Poi cell reference for the named cell
	//
	private CellReference getNamedCell(String name) {
		Name aNamedCell = workbook.getName(name); 
		String ref = aNamedCell.getRefersToFormula();
		return ref.contains(":") ? 
			new CellReference(ref.substring(0, ref.indexOf(":"))) : new CellReference(ref);
	}

	public void writeXlsxFile(String filename) throws Exception {
		try {
			
			System.out.println("\nwriting to file: " + filename);
			FileOutputStream out = new FileOutputStream(new File(filename));
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			throw e;
		}
	}


/* 
	public void handleExcelFile(String cellName) throws Exception {

		String cellLocation = cellName;
		Object cellValue = new Object();

		try {

			// get named cell	
			
			Name aNamedCell = workbook.getName(cellLocation);         
			System.out.println(">>> get named cell " + aNamedCell.getRefersToFormula() + " named-cell-value " + aNamedCell.getNameName());
					
			// named cell props (row / col / val)

			//String ref = aNamedCell.getRefersToFormula().substring(0, aNamedCell.getRefersToFormula().indexOf(":")); //"Sheet1!$D$4";
			String ref = aNamedCell.getRefersToFormula(); 
			CellReference cellReference = new CellReference(ref);
			Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
			Row row = workSheet.getRow(cellReference.getRow());
			int col = cellReference.getCol();
			Cell cell = row.getCell(cellReference.getCol());
			
			System.out.println( ">>> col-row-val " + cellReference.getCol() + "-" + cellReference.getRow() + "-" + cell.getStringCellValue()); 

			// get named col data 

			int startRow = cellReference.getRow();
			Row r = null; 

			do {
				r =  workSheet.getRow(++startRow); 

				if (r != null) {

					Cell c = r.getCell(col); 

					switch (c.getCellType()) {
						case NUMERIC:
							System.out.print(c.getNumericCellValue());
							continue;
						case STRING:
							System.out.print(c.getStringCellValue());
							continue;
						case BLANK:
					    case _NONE:
						case ERROR:
							break;
					} 					
				} 
			} while (r != null);

			//workbook.close();

		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
        return;
	}
		
	
	public CellStyle createWarningColor() {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Courier New");
        font.setBold(false);
        //font.setUnderline(Font.U_SINGLE);
        //font.setColor(org.apache.poi.hssf.util.HSSFColorPredefined.DARK_RED.getIndex());
        style.setFont(font);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER); 
		return style;
	}	
	
	*/
}
