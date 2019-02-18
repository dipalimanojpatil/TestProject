package Com.WriteExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperation {

	XSSFWorkbook wb;
	XSSFRow ro;
	XSSFSheet sheet;
	static WriteOperation wr;
	DataFormatter df = new DataFormatter();
	int rowNo = 0;
	int CellNo = 0;

	public void ReadData(String filenm, String sheetnm) throws IOException {
		FileInputStream fis = new FileInputStream(filenm);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheet(sheetnm);
		int row = sheet.getLastRowNum();
		int rowcount = row + 1;
		System.out.println("Rows : " + rowcount);

		int cells = sheet.getRow(row).getLastCellNum();
		System.out.println("Colunm : " + cells);

		for (int i = 0; i <= rowcount; i++) {
			
			for (int j = 0; j <= cells; j++) {
				String value = df.formatCellValue(sheet.getRow(i).getCell(j));
				System.out.println(value);
				wr.WriteData("D:\\Deepa(Testing)\\ExcelOperations\\WriteDataEmployee.xlsx", "Sheet1", rowNo, CellNo,
						value);
			}
		}

	}

	public void WriteData(String filenm, String sheetnm, int rowno, int colno, String val) throws IOException {

		FileInputStream fis1 = new FileInputStream(filenm);
		XSSFWorkbook wb1 = new XSSFWorkbook(fis1);
		XSSFSheet sheet1= wb.getSheet(sheetnm);
		XSSFRow rows = sheet.createRow(rowno);
		XSSFCell cell = rows.createCell(colno);
		cell.setCellValue(val);

		FileOutputStream fos = new FileOutputStream(filenm);
		wb.write(fos);
	}

	public static void main(String[] args) throws IOException {

		wr = new WriteOperation();
		wr.ReadData("D:\\Deepa(Testing)\\ExcelOperations\\ReadDataEmployee.xlsx", "EmployeeData");

	}

}
