     package excel_test;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class excel_sheet {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
String Path = "C:\\excel_try.xlsx";
FileInputStream file = new FileInputStream(Path); //to read the sheet

//create workbook object 
Workbook  xbook = WorkbookFactory.create(file); //create workbook object XHSFWoorkbok  xlsxHSFWorkbook xls

// create worksheet object 
Sheet sh = xbook.getSheet("sheet1");// with the reference of workbook

// get row 
Row row = sh.getRow(0); //with the reference of sheet

//get cell 
Cell cell =row.getCell(0);//with the reference of row

System.out.println(cell.getStringCellValue()); //to print that particular value

//get last row number 
int rowCount = sh.getLastRowNum();
System.out.println(rowCount);

//get last cell number 
int cellCount = row.getLastCellNum();
System.out.println(cellCount);

for(int i=0; i<=rowCount; i++) {
	Row r1 = sh.getRow(i);
	
	for(int j=0; j<cellCount;j++) {
		cell = r1.getCell(j);
		
		if(cell.getCellType()==CellType.STRING) {
		System.out.println(cell.getStringCellValue());
		}else {
			System.out.println(cell.getNumericCellValue());
		}
	}
}


//String value = WorkbookFactory.create(file).getSheet("sheet1").getRow(1).getCell(1).getStringCellValue();
Thread.sleep(3000);

//System.out.println(value);


	}

}
