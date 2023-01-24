package com.example.excel.stamper;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.io.ByteArrayOutputStream;
import com.lowagie.text.pdf.PdfWriter;
import com.lowagie.text.pdf.PdfReader;
import com.lowagie.text.pdf.PdfStamper;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfImportedPage;

import com.lowagie.text.Element;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class StamperApplication {

/* 	private PdfStamper pdfStamper = null;
	private ByteArrayOutputStream baos = null;

	protected void getPdfTemplate(String template) 
		throws Exception {

		try {

			PdfReader reader = new PdfReader(template+".pdf");
			baos = new ByteArrayOutputStream();
			//pdfStamper = new PdfStamper(reader, new FileOutputStream("instance.pdf"));
			pdfStamper = new PdfStamper(reader, baos);		
	
		} catch (Exception e) {

			throw(e);
		}
	}

	public void removePages(int... pagesToRemove) {
		int pagesTotal = pdfStamper.getReader().getNumberOfPages();
		List<Integer> allPages = new ArrayList<>(pagesTotal);

		for (int i = 1; i <= pagesTotal; i++) {
			allPages.add(i);
		}
		for (int page : pagesToRemove) {
			allPages.remove(new Integer(page));
		}
		pdfStamper.getReader().selectPages(allPages);
	}
		
	protected int getPageCount() {

		return pdfStamper.getReader().getNumberOfPages();
	}
	protected void savePdf(String name) 
		throws Exception {
		
		try {

			pdfStamper.close();
			FileOutputStream fios = new FileOutputStream(name +".pdf");
			baos.writeTo(fios);

		} catch (Exception e) {
			throw e;
		}
	}

	protected void setEncryption(String userPass, String ownerPass) 
		throws Exception {

		try {
		
			pdfStamper.setEncryption(userPass.getBytes(), ownerPass.getBytes(), 
				PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_COPY, PdfWriter.ENCRYPTION_AES_128);

		} catch (Exception e) {
			System.err.println("test "+e);
			throw e;
		}
	}

	protected void insertPage(String page, int position) 
		throws Exception {

		try {

			PdfReader srcReader = new PdfReader(page+".pdf");
			pdfStamper.insertPage(position, srcReader.getPageSize(1));
			PdfImportedPage importedPage = pdfStamper.getImportedPage(srcReader, 1);
			pdfStamper.getUnderContent(position).addTemplate(importedPage, 0, 0);
			srcReader.close();	

		} catch (Exception e) {

			throw(e);
		}	
	}	

	protected void setRandomGrid(int page) 
		throws Exception {

		try {

			BaseFont bf = BaseFont.createFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.EMBEDDED);
			PdfContentByte cb = pdfStamper.getOverContent(page);
			
			cb.beginText();
			cb.setFontAndSize(bf, 5);
			cb.setColorFill(java.awt.Color.blue);

			for (int i=0; i<1500; i++) {

				int x = (int) (Math.random() * 1000);
				int y = (int) (Math.random() * 1000); 

				cb.showTextAligned(PdfContentByte.ALIGN_CENTER, new String("["+x+":"+y+"]"), x , y, 0);
			}

			cb.endText();	

		} catch (Exception e) {

			throw(e);
		}
	}

	protected void setDraft(int page) 
			throws Exception {

		try {

			BaseFont bf = BaseFont.createFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.EMBEDDED);
			int tSize = 36;
			String mark = "DRAFT ";
			int angle = 45;
			com.lowagie.text.Rectangle rect = pdfStamper.getReader().getPageSizeWithRotation(page);
			float height = rect.getHeight() / 2; 
			float width = rect.getWidth() / 2; 
			PdfContentByte cb = pdfStamper.getOverContent(page);

			cb.setColorFill(java.awt.Color.black);
			cb.setFontAndSize(bf, tSize);
			cb.beginText();
			cb.showTextAligned(Element.ALIGN_CENTER, mark, width, height, angle);
			cb.endText();

		} catch (Exception e) {

			throw(e);
		}
	}	*/

	public static void main(String[] args) {
		SpringApplication.run(StamperApplication.class, args);
	}

}
