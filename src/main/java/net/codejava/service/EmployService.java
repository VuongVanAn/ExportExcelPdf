package net.codejava.service;

import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.codejava.model.Employee;

public interface EmployService {
	
	List<Employee> getAll();

	boolean createPdf(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response);

	boolean createExcel(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response);
}
