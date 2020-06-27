package net.codejava.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import net.codejava.model.Employee;
import net.codejava.service.EmployService;

@RestController
public class PdfExcelController {
	
	@Autowired
	private EmployService service;
	
	@Autowired
	private ServletContext context;
	
	@RequestMapping("/")
	public String allEmployees(Model model) {
		List<Employee> employees = service.getAll();
		model.addAttribute("employees", employees);
		return "view/employees";
	}
	
	@GetMapping("/createPdf")
	public void createPdf(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = service.getAll();
		boolean isFlag = service.createPdf(employees, context, request, response);
		if (isFlag) {
			String fullPath = request.getServletContext().getRealPath("/resources/reports/" + "employees" + ".pdf");
			fileDownload(fullPath, response, "employees.pdf");
		}
	}
	
	@GetMapping("/createExcel")
	public void createExcel(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = service.getAll();
		boolean isFlag = service.createExcel(employees, context, request, response);
		if (isFlag) {
			String fullPath = request.getServletContext().getRealPath("/resources/reports/" + "employees" + ".xls");
			fileDownload(fullPath, response, "employees.xls");
		}
	}

	private void fileDownload(String fullPath, HttpServletResponse response, String fileName) {
		File file = new File(fullPath);
		final int BUFFER_SIZE = 4096;
		if (file.exists()) {
			try {
				FileInputStream inputStream = new FileInputStream(file);
				String mimeType = context.getMimeType(fullPath);
				response.setContentType(mimeType);
				response.setHeader("content-disposition", "attachment; filname=" + fileName);
				OutputStream outputStream = response.getOutputStream();
				byte[] buffer = new byte[BUFFER_SIZE];
				int bytesRead = -1;
				while((bytesRead = inputStream.read(buffer)) != -1) {
					outputStream.write(buffer, 0, bytesRead);
				}
				inputStream.close();
				outputStream.close();
				file.delete();
				
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
	}
	
}
