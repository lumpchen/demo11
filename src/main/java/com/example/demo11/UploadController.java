package com.example.demo11;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class UploadController {
	@ResponseBody
	@RequestMapping(value = "/Upload", method = RequestMethod.POST)
	public String upload(@RequestParam("file") MultipartFile files, HttpServletRequest request,
			HttpServletResponse response) {
		
		File workDir = new File(request.getServletContext().getContextPath());
		String s = workDir.getAbsolutePath();
		File upload = new File(workDir.getAbsolutePath(), "upload");
		if (!upload.exists()) {
			upload.mkdirs();
		}
		
		// Client File Name
		String name = files.getOriginalFilename();
		System.out.println("Client File Name = " + name);

		if (name != null && name.length() > 0) {
			try {
				// Create the file at server
				File serverFile = new File(upload, name);

				BufferedOutputStream stream = new BufferedOutputStream(new FileOutputStream(serverFile));
				stream.write(files.getBytes());
				stream.close();

				System.out.println("Write file: " + serverFile);
				
				// 1) Load docx with POI XWPFDocument
				XWPFDocument document = new XWPFDocument(new FileInputStream(serverFile));

				// 2) Convert POI XWPFDocument 2 PDF with iText
				
				File outFile = new File(upload, name + ".pdf");

				OutputStream out = new FileOutputStream(outFile);
				PdfOptions options = PdfOptions.create().fontEncoding("windows-1250");
				PdfConverter.getInstance().convert(document, out, options);

				String path = "\\upload\\" + name + ".pdf";
				return "<a href=\"" + path + "\"" + ">" + "download pdf" + "</a>";
			} catch (Exception e) {
				e.printStackTrace();
				System.out.println("Error Write file: " + name);
			}
		}

		return name + " upload succeeded";
	}
}
