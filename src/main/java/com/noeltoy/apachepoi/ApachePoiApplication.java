package com.noeltoy.apachepoi;

import com.noeltoy.apachepoi.reader.FileReader;
import com.noeltoy.apachepoi.reader.FileReaderDocx;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ApachePoiApplication {

	public static void main(String[] args) {
		SpringApplication.run(ApachePoiApplication.class, args);
		/*FileReader fileReader = new FileReader();
		fileReader.docReader();*/

		FileReaderDocx fileReaderDocx = new FileReaderDocx();
		fileReaderDocx.docxReader();
	}

}
