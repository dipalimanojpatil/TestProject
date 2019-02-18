package Com.ReadExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.record.PageBreakRecord.Break;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadOperationFrom_xlsx {

	public void ReadData(String file, String Sheet) throws IOException {

		int arraydata[][] = null;
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheett = wb.getSheet(Sheet);
		XSSFRow row = sheett.getRow(3);
	
		// No. of Rows
		int rows = sheett.getLastRowNum();
		System.out.println("No of Rows By Index: " + rows);
		int rowcount = rows + 1;
		System.out.println("No Of Original Rows: " + rowcount);

		// No. of Colums
		int colnm = row.getLastCellNum();
		System.out.println("No. of Colunm is: " + colnm);

		// fatch all data
		for (int i = 0; i < rowcount; i++) {
			for (int j = 0; j < colnm; j++) {
				System.out.println(sheett.getRow(i).getCell(j));
				
			}
		}
	}


	
	
	
	public static void main(String[] args) throws IOException {

		ReadOperationFrom_xlsx obj = new ReadOperationFrom_xlsx();
		obj.ReadData("D:\\Deepa(Testing)\\ExcelOperations\\StudentData_xlsx.xlsx", "StudentInfo");

	}

}
