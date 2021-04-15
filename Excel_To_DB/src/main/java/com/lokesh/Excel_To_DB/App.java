package com.lokesh.Excel_To_DB;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.Transaction;
import org.hibernate.cfg.Configuration;
import org.hibernate.service.ServiceRegistry;
import org.hibernate.service.ServiceRegistryBuilder;

public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	
    	Configuration con = new Configuration().configure().addAnnotatedClass(Employee.class);
    	
    	ServiceRegistry reg = new ServiceRegistryBuilder().applySettings(con.getProperties()).buildServiceRegistry();
    	
    	SessionFactory sf = con.buildSessionFactory(reg);
    	
    	Session session = sf.openSession();
    	
    	Transaction tx =session.beginTransaction();
    	
		File file = new File("C:\\poi_files\\emp_data.xlsx");
		
		FileInputStream fis = new FileInputStream(file);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int id_data = (int) sheet.getRow(1).getCell(0).getNumericCellValue();
		System.out.println(id_data);
		

		String designation_data = sheet.getRow(1).getCell(1).getStringCellValue();
		System.out.println(designation_data);
		

		String name_data = sheet.getRow(1).getCell(2).getStringCellValue();
		System.out.println(name_data);
		

		double salary_data = sheet.getRow(1).getCell(3).getNumericCellValue();
		System.out.println(salary_data);
		
		workbook.close();
		fis.close();
		
		
    	
    	
    	Employee jocata = new Employee();
//    	jocata.setEmp_id(102);
//    	jocata.setName("Nagesh");
//    	jocata.setSalary(15000);
//    	jocata.setDesignation("Intern");
    	
    	jocata.setEmp_id(id_data);
    	jocata.setName(name_data);
    	jocata.setSalary(salary_data);
    	jocata.setDesignation(designation_data);
    	
    	session.save(jocata);
    	
    	tx.commit();
    	
    	session.close();
    }
}
