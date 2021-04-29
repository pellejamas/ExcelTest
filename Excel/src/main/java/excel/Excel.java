package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
public class Excel {
	
	public static void main(String[] args) {
		//Create workbook in .xls format
		Workbook workbook = new HSSFWorkbook();
		//For .xslx workbooks use XSSFWoorkbok();
		//Create Sheets
		Sheet sh = workbook.createSheet("Decathlon");
		Sheet sh2 = workbook.createSheet("Heptathlon");
		//Create Rows and Cells for Decathlon ROW 1
		Cell cellA1 = sh.createRow(0).createCell(0);
		Cell cellB1 = sh.createRow(0).createCell(1);
		Cell cellC1 = sh.createRow(0).createCell(2);
		Cell cellD1 = sh.createRow(0).createCell(3);
		Cell cellE1 = sh.createRow(0).createCell(4);
		Cell cellF1 = sh.createRow(0).createCell(5);
		Cell cellG1 = sh.createRow(0).createCell(6);
		Cell cellH1 = sh.createRow(0).createCell(7);
		Cell cellI1 = sh.createRow(0).createCell(8);
		Cell cellJ1 = sh.createRow(0).createCell(9);
		Cell cellK1 = sh.createRow(0).createCell(10);
		Cell cellL1 = sh.createRow(0).createCell(11);
		Cell cellM1 = sh.createRow(0).createCell(12);
		Cell cellN1 = sh.createRow(0).createCell(13);
		Cell cellO1 = sh.createRow(0).createCell(14);
		
		//Add value in a cell for Decathlon ROW 1
		cellA1.setCellValue("Number");
		cellB1.setCellValue("Name");
		cellC1.setCellValue("Event 1");
		cellD1.setCellValue("Event 2");
		cellE1.setCellValue("Event 3");
		cellF1.setCellValue("Event 4");
		cellG1.setCellValue("Event 5");
		cellH1.setCellValue("Day 1 points");
		cellI1.setCellValue("Event 6");
		cellJ1.setCellValue("Event 7");
		cellK1.setCellValue("Event 8");
		cellL1.setCellValue("Event 9");
		cellM1.setCellValue("Event 10");
		cellN1.setCellValue("Day 2 points");
		cellO1.setCellValue("Total points");
		
		//Create Rows and Cells for Decathlon ROW 2
		Cell cellA2 = sh.createRow(1).createCell(0);
		Cell cellB2 = sh.createRow(1).createCell(1);
		Cell cellH2 = sh.createRow(1).createCell(7);
		Cell cellN2 = sh.createRow(1).createCell(13);
		Cell cellO2 = sh.createRow(1).createCell(14);
		//Add value in a cell for Decathlon ROW 2
		cellA2.setCellValue("1:A");
		cellB2.setCellValue("Pär Andersson");
		cellH2.setCellValue("3000");
		cellN2.setCellValue("4000");
		cellO2.setCellFormula("SUM(H2+N2)");
		
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
