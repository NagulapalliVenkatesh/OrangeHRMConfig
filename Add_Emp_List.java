package com.AddEmployeeListConfig;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

public class Add_Emp_List extends Base_Emp_List{
	FileInputStream propertiesFile;

	Properties properties;

	@Test(priority=1,description="opening the OrnageHRM application with valid test data")
	public void applicationLogin() throws IOException
	{
	FileInputStream propertiesFile=new FileInputStream("./configurationProperties/AddEmployeeList.properties");
	properties=new Properties();
	properties.load(propertiesFile);

	By usernameProperty=By.name(properties.getProperty("usernameNameProperty"));
	WebElement usernameElement=driver.findElement(usernameProperty);
	usernameElement.sendKeys("venky");
	By passwordProperty=By.name(properties.getProperty("passwordNameProperty"));
	WebElement passwordElement=driver.findElement(passwordProperty);
	passwordElement.sendKeys("Venky@123");
	By buttonProperty=By.className(properties.getProperty("submitbuttonclassnameProperty"));
	WebElement buttonElement=driver.findElement(buttonProperty);
	buttonElement.click();
	}

	@Test(priority=2,description="Identifying the PIM in the OrnageHRM application")
	public void identifing_PIM()
	{
	By pim_Property=By.linkText(properties.getProperty("PIMlinktextProperty"));
	WebElement pim_Element=driver.findElement(pim_Property);
	Actions mouse=new Actions(driver);
	mouse.moveToElement(pim_Element).build().perform();
	}

	@Test(priority=3,description="Identifying the Employee List in PIM in the OrnageHRM application")
	public void identifying_emp_list()
	{
	By emp_List_Property=By.linkText(properties.getProperty("employeelistlinktextProperty"));
	WebElement emp_list_Element=driver.findElement(emp_List_Property);
	emp_list_Element.click();
	}

	@Test(priority=4,description="Getting the Employee List in PIM in the OrnageHRM application")
	public void getting_Emp_List_Data() throws IOException
	{
	FileInputStream inputtestdata=new FileInputStream("./ExcelSheets/OarangeHRM_Employee_List_config.xlsx");
	XSSFWorkbook workBook= new XSSFWorkbook(inputtestdata);
	XSSFSheet sheetOfWorkBook=workBook.getSheet("testdata");
	By tableProperty=By.xpath(properties.getProperty("webtablebodyxpathProperty"));
	WebElement tableElement=driver.findElement(tableProperty);
	By row_of_table_Property=By.tagName(properties.getProperty("rowoftabletagnameProperty"));
	List<WebElement> row_List=tableElement.findElements(row_of_table_Property);
	for(int rowIndex=0;rowIndex<row_List.size();rowIndex++)
	{
	Row rowofTable=sheetOfWorkBook.createRow(rowIndex);
	WebElement row_Index=row_List.get(rowIndex);
	By cell_of_table_Property=By.tagName(properties.getProperty("celloftabletagnameProperty"));
	List<WebElement> cell_list=row_Index.findElements(cell_of_table_Property);
	for(int cellIndex=0;cellIndex<cell_list.size();cellIndex++)
	{
	Cell cellofrow=rowofTable.createCell(cellIndex);
	WebElement cell_index=cell_list.get(cellIndex);
	String cell_text=cell_index.getText();
	cellofrow.setCellValue(cell_text);
	System.out.print(cell_text+" ");
	}
	System.out.println();
	}
	By welcomeadminProperty=By.linkText(properties.getProperty("welcomeadminlinktextProperty"));
	WebElement welcomeadminElement=driver.findElement(welcomeadminProperty);
	welcomeadminElement.click();
	By logoutProperty=By.linkText(properties.getProperty("logoutlinktextProperty"));
	WebElement logoutElement=driver.findElement(logoutProperty);
	logoutElement.click();
	FileOutputStream outputdata= new FileOutputStream("./ExcelSheets/OarangeHRM_Employee_List_config.xlsx");
	workBook.write(outputdata);
	}
}
