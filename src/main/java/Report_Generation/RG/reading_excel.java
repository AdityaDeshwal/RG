package Report_Generation.RG;
import java.io.File; 
import java.io.FileInputStream; 
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.FormulaEvaluator; 
import org.apache.poi.ss.usermodel.Row; 

 class reading_excel {
	public void read(String fileName) {
		try { 
			  
            // Reading file from local directory 
			System.out.println("File path: " + fileName);
            FileInputStream file = new FileInputStream( 
                new File(fileName)); 
  
            // Create Workbook instance holding reference to 
            XSSFWorkbook workbook = new XSSFWorkbook(file); 
  
            // Get first/desired sheet from the workbook 
            XSSFSheet sheet = workbook.getSheetAt(0); 
  
            // Iterate through each rows one by one 
            Iterator<Row> rowIterator = sheet.iterator(); 
  
            // Till there is an element condition holds true 
            while (rowIterator.hasNext()) { 
  
                Row row = rowIterator.next(); 
  
                // For each row, iterate through all the 
                // columns 
                Iterator<Cell> cellIterator 
                    = row.cellIterator(); 
  
                while (cellIterator.hasNext()) { 
  
                    Cell cell = cellIterator.next(); 
  
//                     Checking the cell type and format 
//                     accordingly 
                    switch (cell.getCellType()) { 
  
                    // Case 1 
                    case NUMERIC: 
                        System.out.print( 
                            cell.getNumericCellValue() 
                            + "t"); 
                        break; 
  
                    // Case 2 
                    case STRING: 
                        System.out.print( 
                            cell.getStringCellValue() 
                            + "t"); 
                        break; 
                    } 
                } 
  
                System.out.println(""); 
            } 
  
            // Closing file output streams 
            file.close(); 
        } 
  
        // Catch block to handle exceptions 
        catch (Exception e) { 
  
            // Display the exception along with line number 
            // using printStackTrace() method 
            e.printStackTrace(); 
        } 
	}
}
