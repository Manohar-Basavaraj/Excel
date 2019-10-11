import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MultipleValues {

	public void readExcel(String filePath, String fileName, String sheetName) throws IOException {

		// Create an object of File class to open xlsx file

		File file = new File(filePath + "\\" + fileName);

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook guru99Workbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			guru99Workbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of HSSFWorkbook class

			guru99Workbook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		org.apache.poi.ss.usermodel.Sheet guru99Sheet = guru99Workbook.getSheetAt(0);

		// Find number of rows in excel file

		int rowCount = guru99Sheet.getLastRowNum() - guru99Sheet.getFirstRowNum();

		// Create a loop over all the rows of excel file to read it

		for (Row row : guru99Sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {

					System.out.println(cell.getStringCellValue());

				} else {

					System.out.println(cell.getDateCellValue());
				}
			}
		}

	}



	// Main function is calling readExcel function to read data from excel file

	public static void main(String... strings) throws IOException {

		// Create an object of ReadGuru99ExcelFile class

		MultipleValues objExcelFile = new MultipleValues();

		// Prepare the path of excel file

		String filePath = System.getProperty("user.dir") + "\\src\\";

		// Call read file method of the class to read data

		objExcelFile.readExcel(filePath, "TestData.xlsx", "Sheet1");

	}

}
