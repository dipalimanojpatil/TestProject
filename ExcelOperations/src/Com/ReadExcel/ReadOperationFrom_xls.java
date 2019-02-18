package Com.ReadExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class ReadOperationFrom_xls {

	public void ReadData(String file, String Sheet) throws IOException {

		int arraydata[][] = null;
		FileInputStream fis = new FileInputStream(file);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheett = wb.getSheet(Sheet);
		HSSFRow row = sheett.getRow(3);
		HSSFCell cell = row.getCell(2);

		String Val = cell.getStringCellValue();
		System.out.println("Value For That Index : " + Val);

		// No. of Rows
		int rows = sheett.getLastRowNum();
		System.out.println("No of Rows By Index: " + rows);
		int rowcount = rows + 1;
		System.out.println("No Of Original Rows: " + rowcount);

		// No. of Colums
		int colnm = row.getLastCellNum();
		System.out.println("No. of Colunm is: " + colnm);

		//Intialized Array
		arraydata = new int[rowcount][colnm];

		// fatch all data
		for (int i = 0; i < rowcount; i++) {
			for (int j = 0; j < colnm; j++) {
				System.out.println(sheett.getRow(i).getCell(j));
				
			}
		}
	}

	public static void main(String[] args) throws IOException {

		ReadOperationFrom_xls obj = new ReadOperationFrom_xls();
		obj.ReadData("D:\\Deepa(Testing)\\ExcelOperations\\StudentData_xls.xls", "StudentInfo");
	}
}
