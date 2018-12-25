package com.example.demo11;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;

import javax.servlet.http.HttpServletRequest;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ConvertController {

	@GetMapping("/convert")
	public String hello(String surl, HttpServletRequest request) {
		File workDir = new File(request.getServletContext().getContextPath());
		String s = workDir.getAbsolutePath();
		
		File uploadDir = new File(workDir.getAbsolutePath() + "\\target\\classes\\static", "download");
		if (!uploadDir.exists()) {
			uploadDir.mkdirs();
		}

		String path = run(surl, uploadDir, request);
		if (path == null) {
			return "Conversion failed.";
		} else {
			return path;
		}
	}

	String run(String surl, File uploadDir, HttpServletRequest request) {
		InputStream is = null;
		try {
			URL url = new URL(surl);
			is = url.openStream();

//			XWPFDocument document = new XWPFDocument(new FileInputStream("C:\\dev\\mine\\springboot\\Insurance.docx"));
			XWPFDocument document = new XWPFDocument(is);

//			File workDir = new File(ClassUtils.getDefaultClassLoader().getResource("").getPath());
			String name = surl.substring(surl.lastIndexOf('/') + 1, surl.lastIndexOf('.'));
			File outFile = new File(uploadDir, name + ".pdf");
			OutputStream out = new FileOutputStream(outFile);
			PdfOptions options = PdfOptions.create().fontEncoding("windows-1250");
			PdfConverter.getInstance().convert(document, out, options);

			String reqUrl = request.getRequestURL().toString();
			String downUrl = reqUrl.substring(0, reqUrl.length() - "/convert".length()) + "/download/" + name + ".pdf";
			return downUrl;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return null;
	}
}