package MavenGroup1.Artifact1;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GFG {
	//Creating Sheets in Excel File in Java using Apache POI

	public static void main(String[] args) throws IOException {
		// Creating Workbook instances
        @SuppressWarnings("resource")
		Workbook wb = new XSSFWorkbook();
 
        // An output stream accepts output bytes and
        // sends them to sink
        OutputStream fileOut = new FileOutputStream("C:\\Users\\asus\\Desktop\\Geeks.xlsx");
 
        // Now creating Sheets using sheet object
        Sheet sheet1 = wb.createSheet("Array");
        Sheet sheet2 = wb.createSheet("String");
        Sheet sheet3 = wb.createSheet("LinkedList");
        Sheet sheet4 = wb.createSheet("Tree");
        Sheet sheet5 = wb.createSheet("Dynamic Programing");
        Sheet sheet6 = wb.createSheet("Puzzles");
 
        // Display message on console for successful
        // execution of program
        System.out.println("Sheets Has been Created successfully");
 
        // Finding number of Sheets present in Workbook
        int numberOfSheets = wb.getNumberOfSheets();
        System.out.println("Total Number of Sheets: " + numberOfSheets);
 
        wb.write(fileOut);

	}

}
