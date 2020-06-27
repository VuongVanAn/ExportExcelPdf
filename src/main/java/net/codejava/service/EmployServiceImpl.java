package net.codejava.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.transaction.Transactional;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import net.codejava.model.Employee;
import net.codejava.repository.EmployRepository;

@Service
@Transactional
public class EmployServiceImpl implements EmployService {
	
	@Autowired
	private EmployRepository repo;

	@Override
	public List<Employee> getAll() {
		return (List<Employee>) repo.findAll();
	}
	
	@Override
	public boolean createPdf(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response) {
		Document document = new Document(PageSize.A4, 15, 15, 45, 30);
		try {
			String filePath = context.getRealPath("/resources/reports");
			File file = new File(filePath);
			boolean exists = new File(filePath).exists();
			if (!exists) {
				new File(filePath).mkdirs();
			}
			PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(file + "/" + "employees" + ".pdf"));
			document.open();
			
			Font mainFont = FontFactory.getFont("Arial", 10, BaseColor.BLACK);
			Paragraph paragraph = new Paragraph("All Employees", mainFont);
			paragraph.setAlignment(Element.ALIGN_CENTER);
			paragraph.setIndentationLeft(50);
			paragraph.setIndentationRight(50);
			paragraph.setSpacingAfter(10);
			document.add(paragraph);
			
			PdfPTable table = new PdfPTable(4);
			table.setWidthPercentage(100);
			table.setSpacingBefore(10f);
			table.setSpacingAfter(10f);
			
			Font tableHeader = FontFactory.getFont("Arial", 10, BaseColor.BLACK);
			Font tableBody = FontFactory.getFont("Arial", 10, BaseColor.BLACK);
			
			float[] columnWidths = {2f, 2f, 2f, 2f};
			table.setWidths(columnWidths);
			
			PdfPCell firstName = new PdfPCell(new Paragraph("First Name", tableHeader));
			firstName.setBorderColor(BaseColor.BLACK);
			firstName.setPaddingLeft(10);
			firstName.setHorizontalAlignment(Element.ALIGN_CENTER);
			firstName.setVerticalAlignment(Element.ALIGN_CENTER);
			firstName.setBackgroundColor(BaseColor.GRAY);
			firstName.setExtraParagraphSpace(5f);
			table.addCell(firstName);
			
			PdfPCell lastName = new PdfPCell(new Paragraph("Last Name", tableHeader));
			lastName.setBorderColor(BaseColor.BLACK);
			lastName.setPaddingLeft(10);
			lastName.setHorizontalAlignment(Element.ALIGN_CENTER);
			lastName.setVerticalAlignment(Element.ALIGN_CENTER);
			lastName.setBackgroundColor(BaseColor.GRAY);
			lastName.setExtraParagraphSpace(5f);
			table.addCell(lastName);
			
			PdfPCell email = new PdfPCell(new Paragraph("Email", tableHeader));
			email.setBorderColor(BaseColor.BLACK);
			email.setPaddingLeft(10);
			email.setHorizontalAlignment(Element.ALIGN_CENTER);
			email.setVerticalAlignment(Element.ALIGN_CENTER);
			email.setBackgroundColor(BaseColor.GRAY);
			email.setExtraParagraphSpace(5f);
			table.addCell(email);
			
			PdfPCell phone = new PdfPCell(new Paragraph("Phone Number", tableHeader));
			phone.setBorderColor(BaseColor.BLACK);
			phone.setPaddingLeft(10);
			phone.setHorizontalAlignment(Element.ALIGN_CENTER);
			phone.setVerticalAlignment(Element.ALIGN_CENTER);
			phone.setBackgroundColor(BaseColor.GRAY);
			phone.setExtraParagraphSpace(5f);
			table.addCell(phone);
			
			for (Employee item : employees) {
				PdfPCell firstNameValue = new PdfPCell(new Paragraph(item.getFirstName(), tableBody));
				firstNameValue.setBorderColor(BaseColor.BLACK);
				firstNameValue.setPaddingLeft(10);
				firstNameValue.setHorizontalAlignment(Element.ALIGN_CENTER);
				firstNameValue.setVerticalAlignment(Element.ALIGN_CENTER);
				firstNameValue.setBackgroundColor(BaseColor.WHITE);
				firstNameValue.setExtraParagraphSpace(5f);
				table.addCell(firstNameValue);
				
				PdfPCell lastNameValue = new PdfPCell(new Paragraph(item.getLastName(), tableBody));
				lastNameValue.setBorderColor(BaseColor.BLACK);
				lastNameValue.setPaddingLeft(10);
				lastNameValue.setHorizontalAlignment(Element.ALIGN_CENTER);
				lastNameValue.setVerticalAlignment(Element.ALIGN_CENTER);
				lastNameValue.setBackgroundColor(BaseColor.WHITE);
				lastNameValue.setExtraParagraphSpace(5f);
				table.addCell(lastNameValue);
				
				PdfPCell emailValue = new PdfPCell(new Paragraph(item.getEmail(), tableBody));
				emailValue.setBorderColor(BaseColor.BLACK);
				emailValue.setPaddingLeft(10);
				emailValue.setHorizontalAlignment(Element.ALIGN_CENTER);
				emailValue.setVerticalAlignment(Element.ALIGN_CENTER);
				emailValue.setBackgroundColor(BaseColor.WHITE);
				emailValue.setExtraParagraphSpace(5f);
				table.addCell(emailValue);
				
				PdfPCell phoneValue = new PdfPCell(new Paragraph(item.getPhoneNumber(), tableBody));
				phoneValue.setBorderColor(BaseColor.BLACK);
				phoneValue.setPaddingLeft(10);
				phoneValue.setHorizontalAlignment(Element.ALIGN_CENTER);
				phoneValue.setVerticalAlignment(Element.ALIGN_CENTER);
				phoneValue.setBackgroundColor(BaseColor.WHITE);
				phoneValue.setExtraParagraphSpace(5f);
				table.addCell(phoneValue);
			}
			
			document.add(table);
			document.close();
			writer.close();
			return true;
			
		} catch(Exception e) {
			return false;
		}
	}

	@Override
	public boolean createExcel(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response) {
		String filePath = context.getRealPath("/resources/reports");
		File file = new File(filePath);
		boolean exists = new File(filePath).exists();
		if (!exists) {
			new File(filePath).mkdirs();
		}
		
		try {
			FileOutputStream outputStream = new FileOutputStream(file + "/" + "employees" + ".xls");
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet workSheet = workbook.createSheet("Employees");
			workSheet.setDefaultColumnWidth(30);
			
			HSSFCellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
			headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			HSSFRow headerRow = workSheet.createRow(0);
			
			HSSFCell firstName = headerRow.createCell(0);
			firstName.setCellValue("First Name");
			firstName.setCellStyle(headerCellStyle);
			
			HSSFCell lastName = headerRow.createCell(1);
			lastName.setCellValue("Last Name");
			lastName.setCellStyle(headerCellStyle);
			
			HSSFCell email = headerRow.createCell(2);
			email.setCellValue("Email");
			email.setCellStyle(headerCellStyle);
			
			HSSFCell phone = headerRow.createCell(3);
			phone.setCellValue("Phone Number");
			phone.setCellStyle(headerCellStyle);
			
			int i = 1;
			for (Employee item : employees) {
				HSSFRow bodyRow = workSheet.createRow(i);
				
				HSSFCellStyle bodyCellStyle = workbook.createCellStyle();
				bodyCellStyle.setFillForegroundColor(IndexedColors.WHITE.index);
				
				HSSFCell firstNameValue = bodyRow.createCell(0);
				firstNameValue.setCellValue(item.getFirstName());
				firstNameValue.setCellStyle(bodyCellStyle);
				
				HSSFCell lastNameValue = bodyRow.createCell(1);
				lastNameValue.setCellValue(item.getLastName());
				lastNameValue.setCellStyle(bodyCellStyle);
				
				HSSFCell emailValue = bodyRow.createCell(2);
				emailValue.setCellValue(item.getEmail());
				emailValue.setCellStyle(bodyCellStyle);
				
				HSSFCell phoneValue = bodyRow.createCell(3);
				phoneValue.setCellValue(item.getPhoneNumber());
				phoneValue.setCellStyle(bodyCellStyle);
				
				i++;			
			}
			
			workbook.write(outputStream);
			workbook.close();
			outputStream.flush();
			outputStream.close();
			return true;
			
		} catch (Exception e) {
			return false;
		}			
	}

}
