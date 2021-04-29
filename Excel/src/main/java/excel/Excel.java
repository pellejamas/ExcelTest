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
		//Create Row and Cell
		Cell cell = sh.createRow(0).createCell(0);	
		//Add value in a cell
		cell.setCellValue("Name");
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
