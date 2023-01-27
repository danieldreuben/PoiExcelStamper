package com.example.excel.stamper.pdfstamper;

import java.util.ArrayList;
import java.util.List;
import java.util.Arrays;
import java.util.Hashtable;

import com.example.excel.stamper.mapper.NameMappingBean;

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
	//String fileLocation = "Test Domestic Standard Upload Template.xlsx";
	Workbook workbook = null;

	public PoiExcelStamper() {
	}

	// @method getJsonFromNamedCols2
	// reads list of named fields from a worksheet (by row)
	// it assumes all subsequent cols in names are in the same worksheet as first  
	// <p>
	// @param names a list of named columns to read
	// @return a hashtable of mapped bean values

	public Hashtable<String, NameMappingBean> getJsonFromNamedCols2(List<String> names) {

		if (names.size() == 0) return null;

		Hashtable<String, NameMappingBean> cols = new Hashtable<String, NameMappingBean>();

		for(String val : names) {
				
			CellReference cellsref = getNamedCell(val);				
			Sheet workSheet = workbook.getSheet(cellsref.getSheetName());
			int startRow = cellsref.getRow();
			Row r = workSheet.getRow(++startRow);
			NameMappingBean nmb = new NameMappingBean(val); 

			for (; r != null; r = workSheet.getRow(++startRow)) {
				Cell c = r.getCell(cellsref.getCol()); 
				
				if (c != null)
				switch (c.getCellType()) {
					case NUMERIC:
						nmb.setValue(c.getNumericCellValue());
						continue;
					case STRING:
						nmb.setValue(c.getStringCellValue());
						continue;
					case BLANK:
					case _NONE:
					case ERROR:
						break;
				} 
			}
			cols.put(val, nmb);   
		}	

		return cols;
	} 

	// @method getJsonFromNamedCols 
	// reads list of named fields from a worksheet (by row)
	// it assumes all subsequent cols in names are in the same worksheet as first  
	// <p>
	// @param names a list of named columns to read
	// @return a hashtable of mapped bean values

	public Hashtable<String, NameMappingBean> getJsonFromNamedCols(List<String> names) {

		if (names.size() == 0) return null;

		Hashtable<String, NameMappingBean> cols = new Hashtable<String, NameMappingBean>();
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
					
					NameMappingBean nmb = (cols.get(val) != null) ? cols.get(val) : new NameMappingBean(val); 
					cols.put(val, nmb);

					if (c != null)
					switch (c.getCellType()) {
						case NUMERIC:
							nmb.setValue(c.getNumericCellValue());
							continue;
						case STRING:
							nmb.setValue(c.getStringCellValue());
							continue;
						case BLANK:
						case _NONE:
						case ERROR:
							break;
					} 			
				}		
			} 
		} while (r != null);

		return cols;
	} 

	// @method writeToNamedCols2 
	// writes list of named beans to a worksheet (by col) 
	// <p>
	// @param names a beans write

	public void writeToNamedCols2(List<NameMappingBean> names) 
		throws Exception {

		try {

			for(NameMappingBean val : names) {	

				CellReference cellReference = getNamedCell(val.getName());
				Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
				Row headerRow = workSheet.getRow(cellReference.getRow());
				int startRow = headerRow.getRowNum()+1;
				int existRows = workSheet.getLastRowNum(); 
				
				for (int index = 0; index < val.getValues().size(); index++) {

					Row r = (startRow < existRows ? workSheet.getRow(startRow++) : workSheet.createRow(startRow++));
					Cell c = r.createCell(cellReference.getCol());
					CellStyle cs = headerRow.getCell(cellReference.getCol()).getCellStyle();  					
					c.setCellStyle(cs);
					c.setCellValue((String) val.getValues().get(index));	
				}
			}

		} catch (Exception e) {
			throw e;
		}
	} 

	// @method getJsonFromNamedCols 
	// reads list of named fields from a worksheet (by row)
	// it assumes all subsequent cols in names are in the same worksheet as first  
	// <p>
	// @param names a list of named columns to read
	// @return a json doc

	public void writeToNamedCols(List<NameMappingBean> names) 
		throws Exception {

		try {

			CellReference cellReference = getNamedCell(names.get(0).getName());
			Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
			Row headerRow = workSheet.getRow(cellReference.getRow());
			//int startRow = cellReference.getRow();
			Row r = null; 
			int startRow = workSheet.getLastRowNum(); 
			int maxrows = names.stream().mapToInt(p-> p.getValues().size()).max().getAsInt();
			//System.out.println(">>> write-cols " + names);

			for (int index = 0; index < maxrows; index++) {

				r = workSheet.createRow(++startRow); // .getRow(++startRow); 

				//if (r != null) {

					for(NameMappingBean val : names) {		

						if (index < val.getValues().size()) {
					
							Name namedCell = workbook.getName(val.getName()); 

							CellReference cellsref = new CellReference(namedCell.getRefersToFormula());
							Cell c = r.createCell(cellsref.getCol());
							CellStyle cs = headerRow.getCell(cellsref.getCol()).getCellStyle();  
							c.setCellStyle(cs);
							c.setCellValue((String) val.getValues().get(index));		
						}			
					}
				//}
			}

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
		
		for(String val : names){
			if (workbook.getName(val) != null)
				present.add(val); 
		}
		return present;
	} 

	// @method getNames 
	// @param names a list of mapping beans to check within the workbook
	// @return a list of names present within the workbook 
	//
	public List<NameMappingBean> getWorkbookNames(List<NameMappingBean> beans) {
		List<NameMappingBean> present = new ArrayList<NameMappingBean>();		
		
		for(NameMappingBean val : beans) {
			if (workbook.getName(val.getName()) != null)
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

	// @method getWorkbookFromFileInput 
	// gets workbook from file location (used here for testing only) 
	// <p>
	public Workbook getWorkbookFromFileInput(String fileLocation) {

		try {
			FileInputStream inputStream = new FileInputStream(new File(fileLocation));
			this.workbook = new XSSFWorkbook(inputStream);
			FormulaEvaluator evaluator = workbook.getCreationHelper()
				.createFormulaEvaluator();

		} catch (Exception e) {
			e.printStackTrace();	
		}
		return workbook;
	}	

	public void writeXlsxFile(String filename) throws Exception {
		try {
			FileOutputStream out = new FileOutputStream(new File(filename));
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			throw e;
		}
	}

	// @method getJsonFromNamedCols 
	// java reflection example dynamically generates mapping class and invokes setters (incomplete) 
	// 
	// @param mappingBeanClass a java bean class with name setters / getters 
	// @return a list of mappingBeanClass rows 
	public List<Object>  getJsonFromNamedCols2(Class mappingBeanClass) {
		List<Object> names = null;

		try {
			java.beans.BeanInfo bi = java.beans.Introspector.getBeanInfo(mappingBeanClass);
			java.beans.PropertyDescriptor[] pds = bi.getPropertyDescriptors();
			names = new ArrayList<Object>();

			Object mapperbean = mappingBeanClass.getConstructor().newInstance();

			for (int i=0; i<pds.length; i++) {

				String propName = pds[i].getName();

				if (propName.compareTo("class") != 0) {
					String setter = "set" + propName.substring(0, 1).toUpperCase() + propName.substring(1);
					java.beans.Statement stmt = new java.beans.Statement(mapperbean, setter, new Object[]{"My Prop Value"});
					stmt.execute();
					//System.out.println(">>setter " + setter);
				}

			}
			names.add(mapperbean);

		} catch (Exception e) {
			e.printStackTrace();
		}
		return names;
	}

	public List<NameMappingBean> getTestNamedMappingBeans(List<String> names, int max) {

		List<NameMappingBean> beans = new ArrayList<NameMappingBean>();

		for(String val : names) {	
			int random = (int) (java.lang.Math.random() * max + 1); 
			List<String> cells = new ArrayList<String>();

			for (int i = 0; i < random; i++) {
				cells.add(val.substring(0, 3)+":"+i);
			}
			beans.add(new NameMappingBean(val, cells));
		}

		return beans;
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
