package com.example.excel.stamper;

import java.util.List;
import java.util.Arrays;


import com.example.excel.stamper.pdfstamper.PoiExcelStamper;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.beans.factory.annotation.Autowired;

import org.junit.jupiter.api.Test;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.beans.Transient;

@SpringBootTest
class StamperApplicationTests {

	@Autowired
    private PoiExcelStamper stamperApp;

	@Test
	void contextLoads() {
	}

	@Test
	public void testReadMulticols() {
		try {
			String[] strArray = {"color","test","label","size","ross","dds","vpn","typeofbuy","department"};
			List<String> names = Arrays.asList(strArray);
			stamperApp.getJsonFromNamedCols(stamperApp.getNames(names));

		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public void testWriteMulticols() {
		try {
			String[] strArray = {"color","style","department","size","vpn","category"};
			List<String> names = Arrays.asList(strArray);
			stamperApp.writeToNamedCols(stamperApp.getNames(names));
			
		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(false);
		}
		assertTrue(true);
	}

	@Test
	public void testNames() {
		try {
			String[] strArray = {"color","test","label","abcd","ross","dds","vpn","$$$","typeofbuy","department"};
			List<String> names = Arrays.asList(strArray);
			System.out.println("\n>>> names present in workbook: " + stamperApp.getNames(names));
		} catch (Exception e) {
			assertTrue(false);
		}
		assertTrue(true);
	}

/* 
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
