package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
public class Excel {
	
	public static void main(String[] args) {
		//Create workbook in .xls format
		Workbook workbook = new HSSFWorkbook();
		//For .xslx workbooks use XSSFWoorkbok();
		//Create Sheet
		Sheet sh = workbook.createSheet("Decathlon");
		Sheet sh2 = workbook.createSheet("Heptathlon");
		//Create Row and Cell for Decathlon
		Cell cell = sh.createRow(0).createCell(0);
		Cell cell1 = sh.createRow(0).createCell(1);
		Cell cell2 = sh.createRow(0).createCell(2);
		Cell cell3 = sh.createRow(0).createCell(3);
		Cell cell4 = sh.createRow(0).createCell(4);
		Cell cell5 = sh.createRow(0).createCell(5);
		Cell cell6 = sh.createRow(0).createCell(6);
		Cell cell7 = sh.createRow(0).createCell(7);
		Cell cell8 = sh.createRow(0).createCell(8);
		Cell cell9 = sh.createRow(0).createCell(9);
		Cell cell10 = sh.createRow(0).createCell(10);
		Cell cell11 = sh.createRow(0).createCell(11);
		Cell cell12 = sh.createRow(0).createCell(12);
		Cell cell13 = sh.createRow(0).createCell(13);
		Cell cell14 = sh.createRow(0).createCell(14);
		
		//Add value in a cell for Decathlon
		cell.setCellValue("Number");
		cell1.setCellValue("Name");
		cell2.setCellValue("Event 1");
		cell3.setCellValue("Event 2");
		cell4.setCellValue("Event 3");
		cell5.setCellValue("Event 4");
		cell6.setCellValue("Event 5");
		cell7.setCellValue("Day 1 points");
		cell8.setCellValue("Event 6");
		cell9.setCellValue("Event 7");
		cell10.setCellValue("Event 8");
		cell11.setCellValue("Event 9");
		cell12.setCellValue("Event 10");
		cell13.setCellValue("Day 2 points");
		cell14.setCellValue("Total points");
		
		//Write the same value of the cell into the console in java
		System.out.println(cell.getRichStringCellValue().toString());
		
		try {
			//Write the output to file
			FileOutputStream output = new FileOutputStream("C:\\Users\\snick\\Desktop\\Deca-HeptathlonScoreboard.xls");
			workbook.write(output);
			output.close();
			workbook.close();
			System.out.println("Excel-file is Completed");
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}
