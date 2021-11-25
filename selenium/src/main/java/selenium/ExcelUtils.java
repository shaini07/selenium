package selenium;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

@Test
public class ExcelUtils {
	// static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public ExcelUtils(String excelPath, String sheetName) {
		try {
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static int getRowCount() {
		int rowCount = 0;
		try {
			rowCount = sheet.getPhysicalNumberOfRows();
		} catch (Exception exp) {
			// TODO Auto-generated catch block
			exp.printStackTrace();
		}
		return rowCount;

	}

	public static String getCellDatastring(int rowNum, int colNum) {
		String CellData = null;
		try {

			CellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
		} catch (Exception e) {
			// TODO: handle exception
		}
		return CellData;
	}

	public static long getCellDataNumber(int rowNum, int colNum) {
		long CellData = 0;
		try {
			CellData = (long) sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
		} catch (Exception exp) {
			// TODO Auto-generated catch block
			exp.printStackTrace();
		}
		return CellData;

	}

	public static int getColCount() {
		int colCount = 0;
		try {
			colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		} catch (Exception exp) {
			// TODO Auto-generated catch block
		}
		return colCount;
	}

}
