package com.example.export_excel_using_template;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

@SpringBootApplication
public class ExportExcelUsingTemplateApplication extends SpringBootServletInitializer {

	@Override
    protected SpringApplicationBuilder configure(SpringApplicationBuilder builder) {
        return builder.sources(ExportExcelUsingTemplateApplication.class);
    }

	public static void main(String[] args) {
		SpringApplication.run(ExportExcelUsingTemplateApplication.class, args);
	}

}
