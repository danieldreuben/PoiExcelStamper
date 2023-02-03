package com.ross.excel.serializer.naming;

import java.util.List;
import java.util.Arrays;
import java.util.ArrayList;
import java.util.Hashtable;

import com.ross.excel.serializer.mapper.NameMappingBean;

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
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


@Component
public class XlsxNamingAdapter  {
	Workbook workbook = null;

	public XlsxNamingAdapter() {
	}

	// @method getFromNamedCols
	// reads list of named fields from a worksheet (by col)
	// <p>
	// @param names a list of named columns to read
	// @return a hashtable of mapped bean values
	//
	public Hashtable<String, NameMappingBean> getFromNamedCols(List<String> names) {

		if (names.size() == 0) return null;
		//List<String> names = getWorkbookNames(cnames);

		Hashtable<String, NameMappingBean> cols = new Hashtable<String, NameMappingBean>();

		names.forEach( (val) -> { 	

			CellReference cellsref = getLastCellInRange(val);				
			Sheet workSheet = workbook.getSheet(cellsref.getSheetName());
			int startRow = cellsref.getRow();
			Row r = workSheet.getRow(++startRow);
			NameMappingBean nmb = new NameMappingBean(val); 

			for (; r != null; r = workSheet.getRow(++startRow)) {
				Cell c = r.getCell(cellsref.getCol()); 
				
				if (c != null)
				switch (c.getCellType()) {
					case NUMERIC:
						nmb.add(c.getNumericCellValue());
						continue;
					case STRING:
						nmb.add(c.getStringCellValue());
						continue;
					case BLANK:
					case _NONE:
					case ERROR:
						break;
				} 
			}
			cols.put(val, nmb);   
		});	

		return cols;
	} 

    // @method writeWithinRange
	// writes named beans into an named range. NameMappingBean name correlates to an excel range
	// @param names a list of mapper beans to write
	//
	private void writeWithinRange(NameMappingBean bean) {
		try {
				AreaReference ar = getAreaReference(bean.getName());
				CellReference[] allCells = ar.getAllReferencedCells();
				int beanval = 0;

				for (CellReference cellRef : allCells) {

					int startRow = cellRef.getRow();
					Sheet workSheet = workbook.getSheet(cellRef.getSheetName());
					int existRows = workSheet.getLastRowNum(); 

					Row r = (startRow < existRows ? 
							workSheet.getRow(startRow++) : workSheet.createRow(startRow++));					
					Cell c = r.createCell(cellRef.getCol()); 	
					CellStyle cs = workSheet.getColumnStyle(cellRef.getCol()) == null ?
							workbook.getCellStyleAt(0) : workSheet.getColumnStyle(cellRef.getCol());		
					c.setCellStyle(cs);
					setCellFromBeanType(c, bean.getValues().get(beanval++));
				}
		} catch (Exception e) {
			throw e;
		}		
	}

	// @method writeRelativeLocation
	// writes named beans relative to a named cell / range (MappingBean name correlates to an excel range). 
	// Range in this case name is used to determine the starting cell position.  
	// positional property.
	// @param names a list of mapper beans to write
	//
	public void writeRelativeLocation(List<NameMappingBean> names) 
	
		throws Exception {

		try {
			names.forEach( (val) -> { 	

				CellReference cellReference = getLastCellInRange(val.getName());
				Sheet workSheet = workbook.getSheet(cellReference.getSheetName());
				int startRow = cellReference.getRow();
				int existRows = workSheet.getLastRowNum(); 
				int max = val.getValues().size();
				
				for (int index = 0; index < max; index++) {

					Row r = (++startRow < existRows ? 
							workSheet.getRow(startRow) : workSheet.createRow(startRow));					
					Cell c = r.createCell(cellReference.getCol()); 	
					CellStyle cs = workSheet.getColumnStyle(cellReference.getCol()) == null ?
							workbook.getCellStyleAt(0) : workSheet.getColumnStyle(cellReference.getCol());		
					c.setCellStyle(cs);
					setCellFromBeanType(c, val.getValues().get(index));
				}
			});
		} catch (Exception e) {
			throw e;
		}
	} 

	// @method setCellWithType
	// sets cell using bean type  
	//	
	private static void setCellFromBeanType(Cell c, Object val) {

		if (val instanceof String)
			c.setCellValue((String) val);
		else if (val instanceof Double)
			c.setCellValue((Double) val);
		else if (val instanceof Integer)
			c.setCellValue((Integer) val);  		
	}


	// @method createLookups
	// method adds a lookup sheet and writes nmbs to cols (iterates a-z cols so max 26 lookups currently)
	//
	public boolean createLookups(String sheetName, Boolean hidden, List<NameMappingBean> beans) {

		try {
			Sheet ws = workbook.createSheet(sheetName);
			char col = 'A';
			
			for (NameMappingBean val : beans) {
				Name name = workbook.createName();
				name.setNameName(val.getName());
				String range = "'"+sheetName+"'!$"+col+"$1:$"+col+"$"+(val.getValues().size()-1);
				//System.out.println("range : " + val.getName() + ": " + range);
				name.setRefersToFormula(range);
				writeWithinRange(val);
				++col;
			}
			workbook.setSheetHidden(workbook.getSheetIndex(sheetName), true);

		} catch (Exception e) {
			e.printStackTrace();
		}
		return true;
	}

	// @method getNamedCellInRange
	// names in workbook may refer to single / muulti-cell, this method returns last cell. 
	// Generally, reads / write is done after last cell in the range (this may  be improved by allowing adjustment)
	// Restricting marker ranges to 10 cells).  
	// @param a workbook named cell / range 
	// @return last cell in valid range  
	//
	private CellReference getLastCellInRange(String name) throws RuntimeException {
		
		CellReference ref = null;

		try {

			Name namedCellIdx = workbook.getName(name);         			
			AreaReference aref = new AreaReference(namedCellIdx.getRefersToFormula(), workbook.getSpreadsheetVersion());         
			ref = aref.getLastCell(); 

			if (aref.isWholeColumnReference() == true || aref.getAllReferencedCells().length > 10
				|| aref.getLastCell().getCol() != aref.getFirstCell().getCol()) 
				throw new Exception("Invalid marker range-name [must be single col range , < 10 cells]");

		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return ref;
	}	

	// @method getAreaReference 
	// An AreaReference provides a read / write container 
	// @param name a name of a range 
	// @return area reference instance 
	//
	private AreaReference getAreaReference(String name) {

		AreaReference aref = null;

		try {
			Name namedCellIdx = workbook.getName(name);         			
			aref = new AreaReference(namedCellIdx.getRefersToFormula(), workbook.getSpreadsheetVersion());
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return aref;		
	}

	// @method getNames 
	// checks name in workbook
	// @param names a list of names to check within the workbook
	// @return a list of names present within the workbook 
	//
	public List<String> getNames(List<String> names) {

		List<String> present = new ArrayList<String>();		

		names.forEach( (val) -> { 
			if (workbook.getName(val) != null)
				present.add(val); 
		});
		return present;
	} 

	// @method getWorkbookNames 
	// checks mapping beans (names) presence in workbook 
	// @param names a list of mapping beans to check within the workbook
	// @return a list of names present within the workbook 
	//
	public List<NameMappingBean> getWorkbookNames(List<NameMappingBean> beans) {

		List<NameMappingBean> present = new ArrayList<NameMappingBean>();		
		
		beans.forEach( (val) -> { 
			if (workbook.getName(val.getName()) != null)
				present.add(val); 
		});
		return present;
	} 

	// @method getNamedCell 
	// @return a cell reference for the named cell
	//
	private CellReference getNamedCell(String name) {

		Name aNamedCell = workbook.getName(name); 
		String ref = aNamedCell.getRefersToFormula();
		//System.out.println(">>> named cell formula: " + ref);
		return ref.contains(":") ? 
			new CellReference(ref.substring(0, ref.indexOf(":"))) : new CellReference(ref);
	}

	// @method getWorkbookFromFileInput 
	// gets workbook from file location (used here for testing only) 
	// 
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


/* 
	

	// @method getReflectFromNamedCols 
	// java reflection example dynamically generates mapping class and invokes setters 
	// @param mappingBeanClass a java bean class with name setters / getters 
	// @return a list of mappingBeanClass rows 
	//	
	public List<Object>  getReflectFromNamedCols(Class mappingBeanClass) {
		List<Object> names = null;

		try {
			java.beans.BeanInfo bi = java.beans.Introspector.getBeanInfo(mappingBeanClass);
			java.beans.PropertyDescriptor[] pds = bi.getPropertyDescriptors();
			names = new ArrayList<Object>();

			Object mapperbean = mappingBeanClass.getConstructor().newInstance();

			for (int i=0; i<pds.length; i++) {

				String propName = pds[i].getName();

				if (propName.compareTo("class") != 0) {
					String setter = "set" + propName.substring(0, 1).toUpperCase() + 
										propName.substring(1);
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



	*/
}
