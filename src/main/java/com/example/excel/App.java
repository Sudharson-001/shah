package com.example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		int rownum = 1;
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(" Employee Details");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("EMP ID");
		row.createCell(1).setCellValue("EMP NAME");
		row.createCell(2).setCellValue("EMP MAIL ID");
		row.createCell(3).setCellValue("EMP AGE");

		EmployeeDetails emp1 = new EmployeeDetails(975310, "Shah", "shahsudharson@gmail.com",25);
		EmployeeDetails emp2 = new EmployeeDetails(787974, "Priyanka", "fakeidpriya@gmail.com",24);
		EmployeeDetails emp3 = new EmployeeDetails(311610, "g.p.muthu", "paperidnara@gmail.com",18);
		ArrayList<EmployeeDetails> arrList = new ArrayList<>();
		arrList.add(emp1);
		arrList.add(emp2);
		arrList.add(emp3);

		for (EmployeeDetails temp : arrList) {
			int cells = 0;
			row = sheet.createRow(rownum++);
			XSSFCell cell = row.createCell(cells++);
			cell.setCellValue(temp.empId);
			XSSFCell cell2 = row.createCell(cells++);
			cell2.setCellValue(temp.empName);
			XSSFCell cell3 = row.createCell(cells++);
			cell3.setCellValue(temp.empMail);
			XSSFCell cell4 = row.createCell(cells++);
			cell4.setCellValue(temp.empAge);

		}

		try {
			FileOutputStream out = new FileOutputStream(new File("E:\\excel\\EmployeeRecord.xlsx"));

			wb.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("program end");

	}

}
