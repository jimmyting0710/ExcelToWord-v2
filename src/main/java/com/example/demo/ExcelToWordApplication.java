package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelToWordApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelToWordApplication.class, args);
		ExcelToWord excelToWord = new ExcelToWord(); 
		try {
			excelToWord.excelToWordStart();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
