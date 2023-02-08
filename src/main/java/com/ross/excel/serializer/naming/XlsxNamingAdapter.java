package com.ross.excel.serializer.naming;

import java.util.Date;
import java.util.List;
import java.util.Arrays;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.function.Consumer;

import com.ross.excel.serializer.mapper.NameMappingBean;
import com.ross.excel.serializer.mapper.NameMappingBean.contentTypes;

import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.AreaReference;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CreationHelper; 

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



@Component
public class XlsxNamingAdapter  {
	Workbook workbook = null;

	public XlsxNamingAdapter() {
	}

	// @method readFromRange
	// reads from a named range. 
	// @param name of range
	// @return Hashtable of NameMappingBean
	//

	public Hashtable<String, NameMappingBean> readFromRange(List<String> names) {

		Hashtable<String, NameMappingBean> cols = new Hashtable<String, NameMappingBean>();

		names.forEach( (val) -> { 	
			cols.put(val, readFromRange(val));
		});

		return cols;
	}
		
	// @method readFromRange
	// reads from a named range. 
	// @param name of range
	// @return NameMappingBean range content 
	//

	public NameMappingBean readFromRange(String name) {

		NameMappingBean nmb = null;

		try {
				AreaReference ar = getAreaReference(name);
				CellReference[] allCells = ar.getAllReferencedCells();
				nmb = new NameMappingBean(name); 

				for (CellReference cellRef : allCells) {

					Sheet workSheet = workbook.getSheet(cellRef.getSheetName());
					Row r = workSheet.getRow(cellRef.getRow());	

					if (r != null) {		
						Cell c = r.getCell(cellRef.getCol()); 
						if (c != null)
							nmb.add(getCellConent(c));
						else 
							nmb.add((String) null);	
					}					
				}
		} catch (Exception e) {

			throw e;
		}		
		return nmb;
	}

	// @method readRelativeRange
	// reads list of named fields from a worksheet (by col)
	// <p>
	// @param names a list of named columns to read
	// @return a hashtable of mapped bean values
	//

	public Hashtable<String, NameMappingBean> readRelativeRange(List<String> names) {

		Hashtable<String, NameMappingBean> cols = new Hashtable<String, NameMappingBean>();

		names.forEach( (val) -> { 	

			CellReference cellsref = getAreaReference(val).getLastCell();				
			Sheet workSheet = workbook.getSheet(cellsref.getSheetName());
			int startRow = cellsref.getRow();
			Row r = workSheet.getRow(++startRow);
			NameMappingBean nmb = new NameMappingBean(val); 

			for (; r != null; r = workSheet.getRow(++startRow)) {
				Cell c = r.getCell(cellsref.getCol()); 
				if (c != null) 
					nmb.add(getCellConent(c));		
				else	
					nmb.add((String) null);			
			}
			cols.put(val, nmb);   
		});	

		return cols;
	} 

    // @method writeInRange
	// writes named beans into an named range. 
	// @param names a list of mapping beans to write
	//

	public void writeInRange(List<NameMappingBean> names) {

		names.forEach( (val) -> { 	
			writeInRange(val);
		});
	}

    // @method writeInRange
	// writes named beans into an named range. NameMappingBean name correlates to an excel range
	// @param names a mapper beans to write
	//

	public void writeInRange(NameMappingBean bean) {

		try {
				AreaReference ar = getAreaReference(bean.getName());
				CellReference[] allCells = ar.getAllReferencedCells();
				int nextval = 0;
				
				for (CellReference cellRef : allCells) {

					int startRow = cellRef.getRow();
					Sheet workSheet = workbook.getSheet(cellRef.getSheetName());
					Row r = (workSheet.getRow(startRow) != null ? 
							workSheet.getRow(startRow++) : workSheet.createRow(startRow++));											
					Cell c = r.createCell(cellRef.getCol()); 	
					CellStyle cs  = getDefaultOrCloneStyle(workSheet, cellRef, bean);	

					setCellFromBeanType(c, cs, (bean.getValues().size() > nextval) ? 
						bean.getValues().get(nextval++) : null);
				}
		} catch (RuntimeException e) {

			throw e;
		}		
	}

	// @method writeRelativeRange
	// writes named beans relative to a named cell / range (MappingBean name correlates to an excel range). 
	// Range in this case name is used to determine the starting cell position.  
	// positional property.
	// @param names a list of mapper beans to write
	//

	public void writeRelativeRange(List<NameMappingBean> names) 
		throws Exception {

		try {
			names.forEach( (val) -> { 	

				CellReference cellReference = getAreaReference(val.getName()).getLastCell();
				Sheet workSheet = workbook.getSheet(cellReference.getSheetName());				
				int startRow = cellReference.getRow()+1;
				int max = val.getValues().size();

				for (int index = 0; index < max; startRow++, index++) {

					Row r = (workSheet.getRow(startRow) != null ? 
							workSheet.getRow(startRow) : workSheet.createRow(startRow));						
					Cell c = r.createCell(cellReference.getCol()); 	
					CellStyle cs  = getDefaultOrCloneStyle(workSheet, cellReference, val);
					setCellFromBeanType(c, cs, val.getValues().get(index));
				}
			});
		} catch (RuntimeException e) {

			throw e;
		}
	} 

	// @method getDefaultOrCloneStyle
	// Determines if cell should apply default or new  
	//	

	private CellStyle getDefaultOrCloneStyle(Sheet workSheet, CellReference cellReference, NameMappingBean val) {

		CellStyle cs = workSheet.getColumnStyle(cellReference.getCol()) == null ?
			workbook.getCellStyleAt(0) : workSheet.getColumnStyle(cellReference.getCol());	
			
		if (val.getContentType() == contentTypes.DATE)	{
			CreationHelper createHelper = workbook.getCreationHelper();  
			CellStyle cellStyle = (XSSFCellStyle) workSheet.getWorkbook().createCellStyle();
			cellStyle.cloneStyleFrom(cs);	
		
			//cellStyle.setDataFormat(  
			//	createHelper.createDataFormat().getFormat("d/m/yy"));  		
			cs = cellStyle;						
		}
		return cs;
	}

	// @method setCellWithType
	// sets cell type from bean value type  
	//	

	private void setCellFromBeanType(Cell c, CellStyle cs, Object val) {

		if (val == null) 
			c.setBlank();
		else if (val instanceof String)
			c.setCellValue((String) val);
		else if (val instanceof Double)
			c.setCellValue((Double) val);
		else if (val instanceof Integer)
			c.setCellValue((Integer) val);  	
		else if (val instanceof java.util.Date) {
			c.setCellValue((java.util.Date) val);	
		}  				
		c.setCellStyle(cs);
	}

	// @method getCellConent
	// gets typed cell content   
	//	

	private static Object getCellConent(Cell c) {
		Object res = null;

		switch (c.getCellType()) {
			case NUMERIC:
				res = c.getNumericCellValue();
				/*if(DateUtil.isCellDateFormatted(c)) {
					res = c.getDateCellValue();
				}	*/				
				break;
			case STRING:
				res = c.getStringCellValue();
				break;
			case BLANK:
			case _NONE:				
			case ERROR:
				res = "";						
			}
		return res;
	}	

	// @method createLookups
	// method adds a lookup sheet and writes nmbs to cols 
	// iterates a-z cols so max 26 lookups currently
	//

	public boolean addLookupsFromBeans(String sheetName, Boolean hidden, List<NameMappingBean> beans) {

		try {
			Sheet ws = workbook.createSheet(sheetName);
			char col = 'A';
			
			for (NameMappingBean val : beans) {

				Name name = workbook.createName();
				name.setNameName(val.getName());
				int numitems = val.getValues().size();
				String range = String.format("'%s'!$%s$%d:$%s$%d", sheetName, col, 1, col, numitems);
				//System.out.println("range : " + val.getName() + ": " + range);
				name.setRefersToFormula(range);
				writeInRange(val);
				++col;
			}
			workbook.setSheetHidden(workbook.getSheetIndex(sheetName), true);

		} catch (RuntimeException e) {
			throw e; // e.printStackTrace();
		}
		return true;
	}

	// @method checkMarkerRange
	// This method restricts marker type ranges to 10 cells.	
	// Relative read/write is applied after last cell in a range (may be improved by allowing adjustment)  
	// @param a workbook named cell / range 
	// @return last cell in valid range  
	//

	private Boolean checkMarkerRange (AreaReference aref)  {
		
		if (aref.isWholeColumnReference() == true || aref.getAllReferencedCells().length > 10
			|| aref.getLastCell().getCol() != aref.getFirstCell().getCol()) 
				return false;
		return true;
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
			//throw new RuntimeException(e);
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

	// @method getTestMappingBeans 
	// Generates & fills randomly typed mapping beans 
	// @param list of bean names to generate
	// @param max beans to generate
	// @return collection of randomly generated mapping beans

	public static List<NameMappingBean> getTestMappingBeans(List<String> names, int max) {

		List<NameMappingBean> beans = new ArrayList<NameMappingBean>();
		
		names.forEach( (val) -> { 
			beans.add(new NameMappingBean(val));
		});

		Consumer<NameMappingBean> methodbean = (n) -> { 

			List<contentTypes> includes = 
				Arrays.asList(contentTypes.NUMBER, contentTypes.DATE, contentTypes.STRING);
			final NameMappingBean.contentTypes random = NameMappingBean.contentTypes.getRandom(includes);
			final Integer maximum = Integer.valueOf((int) ((Math.random()) * max + 1));

			for (int i = 0; i < maximum; i++) {
				
				switch (random) {
					case NUMBER:							
						n.add((Math.random()) * max);
						break;
					case DATE:					
						n.add(new java.util.Date());
						break;						
					case STRING:					
						n.add(n.getName().substring(0,3) + ":" + i);	
					case MIXED:	// 
					case EMPTY: // EMPTY MIXED excluded
				}							

			}
			//System.out.println(String.format("%s-%s : %s",n.getName(), random.ordinal(), n.getValues()));			
		};
		beans.forEach(methodbean);
		return beans;
	} 	

	// @method getWorkbookFromFile
	// gets workbook from file location 
	// 

	public Workbook getWorkbookFromFile(String fileLocation) 
		throws Exception {

		try {
			FileInputStream inputStream = new FileInputStream(new File(fileLocation));
			this.workbook = new XSSFWorkbook(inputStream);
			FormulaEvaluator evaluator = workbook.getCreationHelper()
				.createFormulaEvaluator();

		} catch (Exception e) {
			throw e; //e.printStackTrace();	
		}
		return workbook;
	}	

	// @method writeWorkbookToFile
	// writes workbook to file  
	// 

	public void writeWorkbookToFile(String filename) throws Exception {

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
