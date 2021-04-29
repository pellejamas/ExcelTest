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
		
		try {
			//Write the output to file
			FileOutputStream output = new FileOutputStream("C:\\Users\\snick\\Desktop\\Deca-HeptathlonScoreboard.xls");
			workbook.write(output);
			output.close();
			workbook.close();
			System.out.println("Completed");
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}
