package ToolsQa.DemoMavenEclipseProject;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.record.chart.DataFormatRecord;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Excel {

	DataFormatter formatter = new DataFormatter();

	@Test
	public void getExcel() throws IOException {
		FileInputStream fis = new FileInputStream("C:\\Users\\DELL\\Documents\\excelDriven.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
//	private int rowCount;
		Object data[][] = new Object[rowCount - 1][colCount];
		for (int i = 0; i < rowCount - 1; i++) {
			// System.out.println(" outer loop starts");
			row = sheet.getRow(i+1);
			// System.out.println(" outer loop ends");
			for (int j = 0; j < colCount; j++) {
				System.out.println(row.getCell(j));
				// XSSFCell cell = row.getCell(j);
				// data[i][j] = formatter.formatCellValue(cell);
			}
		}
		// return data;
	}
}
