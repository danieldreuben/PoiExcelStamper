package com.ross.excel.serializer;

import java.util.List;
import java.util.Map;
import java.util.Arrays;

import java.util.ArrayList;
import java.util.stream.Collectors;

import com.ross.excel.serializer.mapper.NameMappingBean;
import com.ross.excel.serializer.mapper.MappingElement;

import com.ross.excel.serializer.naming.XlsxNamingAdapter;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.beans.factory.annotation.Autowired;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Order;
import org.springframework.core.Ordered;

import org.junit.jupiter.api.MethodOrderer.OrderAnnotation;
import org.junit.jupiter.api.TestMethodOrder;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;


@SpringBootTest
@TestMethodOrder(OrderAnnotation.class)
class StamperApplicationTests {

	String template = "Domestic Standard Upload Template";
	String fileLocation = String.format("%s.xlsx", template); 
	String writeFileLocation = String.format("W-%s", fileLocation);
	String readFileLocation = String.format("W-%s", fileLocation);
	//String readFileLocation = "gaps.xlsx";
	String lookupFileLocation = String.format("W-%s", fileLocation);

	String[] mixedArray = {"color","style","department","size","typeofbuy","material",
		"weight","cornertest", "$$$","vpn","supplierid","category","vendorstyledescription",
			"label","class","ponumber"
	};
	List<String> mixofnames = Arrays.asList(mixedArray);	

	@Autowired
    private XlsxNamingAdapter stamperApp;

	@Test
	void contextLoads() {
	}

	// writes set of mapping beans to named ranges   
	//
	
	@Test
	@Order(1) 
	public void testWriteRelativeRange() {
		try {
			
			System.out.println(">>> Writing test beans..");
			List<NameMappingBean> beans = XlsxNamingAdapter.getTestMappingBeans(mixofnames, 100);
			stamperApp.getWorkbookFromFile(fileLocation);
			stamperApp.writeRelativeRange(stamperApp.getWorkbookNames(beans));
			stamperApp.writeWorkbookToFile(writeFileLocation);
			
		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);
	}
	// wries bean content within a range 
	//

	@Test
	@Order(2)
	public void testWriteInRange() {
		try {
			
			System.out.println(">>> write to named range");		
			stamperApp.getWorkbookFromFile(writeFileLocation);
			String[] strArray = {"testreadrange"};
			List<String> names = Arrays.asList(strArray);			
			List<NameMappingBean> beans = XlsxNamingAdapter.getTestMappingBeans(names, 8);		

			stamperApp.writeInRange(beans);
			stamperApp.writeWorkbookToFile(writeFileLocation);			

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	// creates new sheet and writes lookup data by name
	//

	@Test
	@Order(3) 
	public void testWriteLookups() {
		try {		
			
			stamperApp.getWorkbookFromFile(writeFileLocation);
			List<NameMappingBean> beans = getTestLookups();
			stamperApp.addLookupsFromBeans("domestic upload tmpt lookups", true, beans);
			stamperApp.writeWorkbookToFile(lookupFileLocation);
			
		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	// reads spreadsheet content from within a ranges 
	//

	@Test
	@Order(10)
	public void testReadFromRange() {
		try {			
			stamperApp.getWorkbookFromFile(readFileLocation);
			NameMappingBean nmb = stamperApp.readFromRange("testreadrange");
			System.out.println(String.format(">>>read from range\n%s", nmb));

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	// reads spreadsheet content from set of ranges into mapping bean
	//

	@Test
	@Order(Ordered.LOWEST_PRECEDENCE)
	public void testReadRelativeRange() {
		try {

			System.out.println(">>> read-relative named range..");	
			stamperApp.getWorkbookFromFile(readFileLocation);
			List<NameMappingBean> nmbs = 
				stamperApp.readRelativeRange(stamperApp.getNames(mixofnames));			
			nmbs.stream().limit(5).forEach(e-> System.out.println(
				StamperApplicationTests.trimOutput(e.getValues().toString(),200)));

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	// trims longer string outputs  
	//
	public static String trimOutput(String s, int max) {
		return (s.length() > max ? String.format("%s...",s.substring(0,max)) : s);
	}

	// @method testNameBeansToRows
	// test convert columnsush  to rows
	//

	@Test
	@Order(Ordered.LOWEST_PRECEDENCE)
	public void testNameBeansToRows() {
		try {
			System.out.println(">>> Mapping Beans To Rows..");	
			List<NameMappingBean> nmbs = XlsxNamingAdapter.getTestMappingBeans(mixofnames, 10);
			//List<MappingElement> elements = MappingElement.getInRows(nmbs);	
			List<MappingElement> elements = nmbs.stream()
				.flatMap(p -> p.getValues().stream())
				.collect(Collectors.toList());
			Map<String, List<MappingElement>> elementmap = elements.stream()
				.collect(Collectors.groupingBy(MappingElement::getLabel));	
			elementmap.entrySet().stream().sorted(Map.Entry.comparingByKey()).limit(7)
				.forEach(e-> System.out.println(StamperApplicationTests.trimOutput(e.toString(),250)));
		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);		
	}

	// queries workbook for defined names

	@Test
	@Order(5)
	public void testNames() {

		try {
			System.out.println(">>> names present in workbook: " + 
				stamperApp.getNames(mixofnames));

		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);
	}


	public List<NameMappingBean> getTestLookups() {
		List<NameMappingBean> beans = new ArrayList<NameMappingBean>();

		List<String> names = Arrays.asList(new String[]{"red","white","blue","pink","orange","black","green"});
		NameMappingBean cl = new NameMappingBean("color_lookup", names);
		beans.add(cl);
		
		List<String> suppliers = Arrays.asList(new String[]{"201650 - Enchante Accessories, Inc","326461 - NIKE INC", 
			"326852 - NIKE SHOES", "43391058 - White Mountain", "496464 - POLO BY RALPH LAUREN HOSIERY","43432102 - Conair LLC",
			"43391270 - ADIDAS ", "43423298 - UNDER ARMOUR","43405653 - Revman International Inc.", "43397382 - POLO RALPH LAUREN"});
		NameMappingBean supps = new NameMappingBean("supplier_lookup", suppliers);
		beans.add(supps);
		
		List<String> departments = Arrays.asList(new String[]{"LADIES HOSIERY","MISSY HVYWT","PLUS MS LTWT SLPWR","Mens Slippers",
			"Boys Shoes","BEAUTY","FRAGRANCES","WELLNESS","Ladies Athletic","WINDOW TREATMENTS"});
		NameMappingBean depts = new NameMappingBean("department_lookup", departments);
		beans.add(depts);

		List<String> labels = Arrays.asList(new String[]{"Lauren Ralph Lauren","Polo Ralph Lauren","Ralph Lauren","Nike Golf","Nike Swim"," Nike Air"," Jordan/Nike Air",
			"Home Essentials"," Madden"," Madden Girl"});
		NameMappingBean label = new NameMappingBean("label_lookup", labels);
		beans.add(label);

		List<String> sizes = Arrays.asList(new String[]{"No Size","6.75","10",".5OZ","0-12M","15mm","16\"-18\"","XS","L","37/38"});
		NameMappingBean size = new NameMappingBean("size_lookup", sizes);
		beans.add(size);

		List<String> typeofbuys = Arrays.asList(new String[]{"Upfront","Closeout","Piece Goods​"});
		NameMappingBean typeofbuy = new NameMappingBean("typeofbuy_lookup", typeofbuys);
		beans.add(typeofbuy);

		List<Double> test_doubles = Arrays.asList(new Double[]{0.021,100.3,9.275,100000.734});
		NameMappingBean test_double = new NameMappingBean("test_doubles_lookup",test_doubles);
		beans.add(test_double);		

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
