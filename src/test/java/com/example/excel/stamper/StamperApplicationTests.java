package com.example.excel.stamper;

import java.util.List;
import java.util.Arrays;

import com.example.excel.stamper.pdfstamper.PoiExcelStamper;
import com.example.excel.stamper.mapper.DomesticStandardUploadTemplate;

import com.example.excel.stamper.mapper.NameMappingBean;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.beans.factory.annotation.Autowired;

import org.junit.jupiter.api.Test;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.beans.Transient;

@SpringBootTest
class StamperApplicationTests {
	String fileLocation = "Domestic Standard Upload Template.xlsx";	
	String writeFileLocation = "X-" + fileLocation;

	@Autowired
    private PoiExcelStamper stamperApp;

	@Test
	void contextLoads() {
	}

	
	@Test
	public void testWriteMulticols() {
		try {
			String[] strArray = {"color","style","department","size","typeofbuy","material",
									"vpn","supplierid","category","vendorstyledescription","label","class","ponumber"};
			List<String> names = Arrays.asList(strArray);
			List<NameMappingBean> beans = stamperApp.getTestNamedMappingBeans(names, 50);

			stamperApp.getWorkbookFromFileInput(fileLocation);
			stamperApp.writeToNamedCols2(stamperApp.getWorkbookNames(beans));
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
			stamperApp.getWorkbookFromFileInput(writeFileLocation);
			String[] strArray = {"color","test","label","size","ross","dds","vpn","typeofbuy","department"};
			List<String> names = Arrays.asList(strArray);
			java.util.Hashtable<String, NameMappingBean> nmb = stamperApp.getJsonFromNamedCols2(stamperApp.getNames(names));

			for(String val : strArray) 
				if (nmb.get(val) != null) 
					System.out.println("results : " + nmb.get(val));

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

	// reflection example works (test data current), needs mapping to xlsx read but unlikely runtime efficient
/* 
	@Test
	public void testReadXlsReflectionBeans() {
		try {
			List<Object> domesticUploadList = 
				stamperApp.getJsonFromNamedCols2(DomesticStandardUploadTemplate.class);
			DomesticStandardUploadTemplate domesticUploadBean = 
				(DomesticStandardUploadTemplate) domesticUploadList.get(0);

			//System.out.println("dut color = " + domesticUploadBean.getColor());

		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	@Test
	public void testMappingBean() {

		System.out.println(">>> testMappingBean\n");
		DomesticStandardUploadTemplate domesticUploadTemplate = new DomesticStandardUploadTemplate();

		try {

			Object o = domesticUploadTemplate;
			java.beans.BeanInfo bi = java.beans.Introspector.getBeanInfo(DomesticStandardUploadTemplate.class);
			java.beans.PropertyDescriptor[] pds = bi.getPropertyDescriptors();
			//System.out.println("pds : " + pds);

			for (int i=0; i<pds.length; i++) {

				String propName = pds[i].getName();

				if (propName.compareTo("class") != 0) {
					String getter = "get" + propName.substring(0, 1).toUpperCase() + propName.substring(1);
					java.beans.Expression expr = new java.beans.Expression(o, getter, new Object[0]);
					expr.execute();
					String s = (String)expr.getValue();
					System.out.print(propName + ":" + s + " "); 
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	@Test
	public void testWriteChanges() {
		try {
			stamperApp.writeXlsxFile("Modified.xlsx");

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public void testExcel() {
		
		try {
			stamperApp.handleExcelFile("color");

		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);
	} */	
}
