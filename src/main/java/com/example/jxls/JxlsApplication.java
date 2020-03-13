package com.example.jxls;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import common.utils.JxlsUtils;
import jxls.demo.guide.ObjectCollectionDemo;

@Controller
@SpringBootApplication
public class JxlsApplication {

	public static void main(String[] args) {
		SpringApplication.run(JxlsApplication.class, args);
	}
	@RequestMapping("testJxls")
	public void testJxls(HttpServletResponse response) throws ParseException, FileNotFoundException, IOException{
		ByteArrayOutputStream os = new ByteArrayOutputStream();
        Map<String , Object> model=new HashMap<String , Object>();
        model.put("employees", ObjectCollectionDemo.generateSampleEmployeeData());
        model.put("nowdate", new Date());
        ClassPathResource templateRes = new ClassPathResource("jxls-template/object_collection_template.xlsx");
        JxlsUtils.exportExcel(templateRes.getInputStream(), os, model);
        download(os.toByteArray(), "test-jxls.xlsx", response);
        os.close();
	}
	/**
	 * 文件下载，需要一个文件的InputStream
	 * @param inStream
	 * @param response
	 * @throws IOException 
	 */
	public static void download(byte[] data, String fileName,HttpServletResponse response) throws IOException{
		response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(),"ISO8859-1"));
		response.addHeader("Content-Length", "" + data.length);
		IOUtils.write(data, response.getOutputStream());
	}
	
}
