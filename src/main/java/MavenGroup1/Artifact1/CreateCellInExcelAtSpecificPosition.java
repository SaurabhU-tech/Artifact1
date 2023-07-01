package MavenGroup1.Artifact1;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateCellInExcelAtSpecificPosition {
	//Creating a Cell at specific position in Excel file using Java
	//https://www.geeksforgeeks.org/creating-a-cell-at-specific-position-in-excel-file-using-java/?ref=lbp

	public static void main(String[] args) throws IOException {
		// Creating a workbook instances
        @SuppressWarnings("resource")
		Workbook wb = new XSSFWorkbook();
 
        // Creating output file
        OutputStream os = new FileOutputStream("C:\\Users\\asus\\Desktop\\CompanyData.xlsx");
 
        // Creating a sheet using predefined class
        // provided by Apache POI
        Sheet sheet = wb.createSheet("Company Preparation");
 
        // Creating a row at specific position
        // using predefined class provided by Apache POI
 
        // Specific row number
        Row row = sheet.createRow(1);
 
        // Specific cell number
        Cell cell = row.createCell(1);
 
        // putting value at specific position
        cell.setCellValue("Geeks");
 
        // Finding index value of row and column of give
        // cell
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
 
        // Writing the content to Workbook
        wb.write(os);
 
        // Printing the row and column index of cell created
        System.out.println("Given cell is created at "
                           + "(" + rowIndex + ","
                           + columnIndex + ")");

	}

}
