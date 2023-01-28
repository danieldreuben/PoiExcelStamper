package com.ross.excel.serializer;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.io.ByteArrayOutputStream;

import com.ross.excel.serializer.naming.NamesSerializer;
import com.ross.excel.serializer.mapper.DomesticStandardUploadTemplate;
import com.ross.excel.serializer.mapper.NameMappingBean;

import com.lowagie.text.Element;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
//@ComponentScan(basePackages = {"controllers","domain"})
public class StamperApplication {

	public static void main(String[] args) {
		SpringApplication.run(StamperApplication.class, args);
	}

}
