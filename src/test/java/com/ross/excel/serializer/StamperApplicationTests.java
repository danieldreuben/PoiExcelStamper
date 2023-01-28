package com.ross.excel.serializer;

import java.util.List;
import java.util.Arrays;
import java.util.ArrayList;

import com.ross.excel.serializer.StamperApplication;
import com.ross.excel.serializer.naming.NamesSerializer;
import com.ross.excel.serializer.mapper.DomesticStandardUploadTemplate;
import com.ross.excel.serializer.mapper.NameMappingBean;

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

	@Autowired
    private NamesSerializer stamperApp;

	@Test
	void contextLoads() {
	}
	
	@Test
	public void testWriteMulticols() {
		try {
			String[] strArray = {"color","style","department","size","typeofbuy","material",
									"vpn","supplierid","category","vendorstyledescription","label","class","ponumber"};
			List<String> names = Arrays.asList(strArray);
			List<NameMappingBean> beans = getTestNamedMappingBeans(names, 50);
			stamperApp.getWorkbookFromFileInput(fileLocation);
			stamperApp.writeToNamedCols(stamperApp.getWorkbookNames(beans));
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
			String[] strArray = {"color","test","label","size","ross","dds","vpn","typeofbuy","department","category","ponumber"};
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
