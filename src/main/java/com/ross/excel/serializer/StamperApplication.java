package com.ross.excel.serializer;


import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.beans.factory.annotation.Autowired;

import com.ross.excel.serializer.naming.XlsxNamingAdapter;

@SpringBootApplication
//@ComponentScan(basePackages = {"controllers","domain"})
public class StamperApplication {
		

	@Autowired
    private XlsxNamingAdapter stamperApp;	

	public static void main(String[] args) {
		SpringApplication.run(StamperApplication.class, args);
	}


}
