package com.ross.excel.serializer;

import java.util.List;
import java.util.Arrays;
import java.util.ArrayList;

import com.ross.excel.serializer.StamperApplication;
import com.ross.excel.serializer.naming.ExcelNamesAdapter;
import com.ross.excel.serializer.mapper.NameMappingBean;
//import com.ross.excel.serializer.mapper.DomesticStandardUploadTemplate;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.beans.factory.annotation.Autowired;

import org.junit.jupiter.api.Test;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.beans.Transient;

@SpringBootTest
class StamperApplicationTests {
	String fileLocation = "Domestic Standard Upload.xlsx";	
	String writeFileLocation = "X-" + fileLocation;
	String readFileLocation = "Y-" + fileLocation;
	String lookupFileLocation = "L-" + fileLocation; 

	@Autowired
    private ExcelNamesAdapter stamperApp;

	@Test
	void contextLoads() {
	}
	
	@Test
	public void testWriteMulticols() {
		try {
			String[] strArray = {"color","style","department","size","typeofbuy","material","weight",
									"vpn","supplierid","category","vendorstyledescription","label","class","ponumber"};
			List<String> names = Arrays.asList(strArray);
			List<NameMappingBean> beans = getTestNamedMappingBeans(names, 50);
			stamperApp.getWorkbookFromFileInput(fileLocation);
			stamperApp.writeRelativeLocation(stamperApp.getWorkbookNames(beans));
			stamperApp.writeXlsxFile(writeFileLocation);
			
		} catch (Exception e) {
			//e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public void testReadMulticols() {
		try {
			stamperApp.getWorkbookFromFileInput(readFileLocation);
			String[] strArray = {"color","test","label","size","ross","dds","vpn","weight","typeofbuy","department","category","ponumber"};
			List<String> names = Arrays.asList(strArray);
			java.util.Hashtable<String, NameMappingBean> nmb = stamperApp.getFromNamedCols(stamperApp.getNames(names));

			for(String val : strArray) 
				if (nmb.get(val) != null) 
					System.out.println("results : " + nmb.get(val));

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public void testWriteLookups() {
		try {
			List<NameMappingBean> beans = new ArrayList<NameMappingBean>();

			List<String> names = Arrays.asList(new String[]{"red","white","blue","pink","orange","black","green"});
			NameMappingBean cl = new NameMappingBean("color_lookup", names);
			beans.add(cl);
			List<String> suppliers = Arrays.asList(new String[]{"201650 - Enchante Accessories, Inc","326461 - NIKE INC", 
				"326852 - NIKE SHOES", "43391058 - White Mountain", "496464 - POLO BY RALPH LAUREN HOSIERY","43432102 - Conair LLC ",
				"43391270 - ADIDAS ", "43423298 - UNDER ARMOUR","43405653 - Revman International Inc.", "43397382 - POLO RALPH LAUREN"});
			NameMappingBean supps = new NameMappingBean("supplier_lookup", suppliers);
			beans.add(supps);
			List<String> departments = Arrays.asList(new String[]{"LADIES HOSIERY"," MISSY HVYWT"," PLUS MS LTWT SLPWR"," Mens Slippers",
				"Boys Shoes","BEAUTY","FRAGRANCES","WELLNESS","Ladies Athletic","WINDOW TREATMENTS "});
			NameMappingBean depts = new NameMappingBean("department_lookup", departments);
			beans.add(depts);

			List<String> labels = Arrays.asList(new String[]{"Lauren Ralph Lauren","Polo Ralph Lauren","Ralph Lauren","Nike Golf"," Nike Swim"," Nike Air"," Jordan/Nike Air",
				" Home Essentials"," Madden"," Madden Girl"});
			NameMappingBean label = new NameMappingBean("label_lookup", labels);
			beans.add(label);

			List<String> sizes = Arrays.asList(new String[]{"No Size","6.75","10",".5OZ","0-12M","15mm","16\"-18\"","XS","L","37/38"});
			NameMappingBean size = new NameMappingBean("size_lookup", sizes);
			beans.add(size);

			List<String> typeofbuys = Arrays.asList(new String[]{"Upfront","Closeout","Piece Goods​"});
			NameMappingBean typeofbuy = new NameMappingBean("typeofbuy_lookup", typeofbuys);
			beans.add(typeofbuy);

			stamperApp.getWorkbookFromFileInput(fileLocation);
			stamperApp.createLookups("domestic upload tmpt lookups", true, beans);
			stamperApp.writeXlsxFile(lookupFileLocation);
			
		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	// queries workbook for defined names

	@Test
	public void testNames() {
		try {
			String[] strArray = {"color","test","label","abcd","ross","dds","vpn","$$$","typeofbuy","department","supplierid","category","vendorstyledescription"};
			List<String> names = Arrays.asList(strArray);
			System.out.println("\n>>> names present in workbook: " + stamperApp.getNames(names));
		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);
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

	// reflection example works, needs mapping to xlsx read but unlikely runtime efficient
/* 
	@Test
	public void testReadXlsReflectionBeans() {
		try {
			List<Object> domesticUploadList = 
				stamperApp.getReflectFromNamedCols(DomesticStandardUploadTemplate.class);
			DomesticStandardUploadTemplate domesticUploadBean = 
				(DomesticStandardUploadTemplate) domesticUploadList.get(0);

			//System.out.println("dut color = " + domesticUploadBean.getColor());

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

*/	
}
