package Report_Generation.RG;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;


public class PopulateSheet {
	 public static void populateSheet(Workbook workbook, String sheetName, List<List<Object>> data, String fileName, String test_code) {
	        Sheet sheet = workbook.createSheet(sheetName);
	        
	        int rowNum = 0;
	        for (List<Object> rowData : data) {
	            Row row = sheet.createRow(rowNum++);
	            int colNum = 0;
	            for (Object field : rowData) {
	                Cell cell = row.createCell(colNum++);
	                if (field instanceof String) {
	                    cell.setCellValue((String) field);
	                } else if (field instanceof Integer) {
	                    cell.setCellValue((Integer) field);
	                } else if (field instanceof Double) {
	                    cell.setCellValue((Double) field);
	                }
	            }
	        }
	        
	        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
	            workbook.write(outputStream);
	            saveAsJson(data, "C:\\Users\\adity\\Downloads\\"+sheetName+"_"+test_code);
	            System.out.println("Populated and JSON formed");
	            outputStream.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	 public static void saveAsJson(List<List<Object>> data, String jsonFileName) {
	        ObjectMapper objectMapper = new ObjectMapper();
	        try (FileWriter fileWriter = new FileWriter(jsonFileName)) {
	            objectMapper.writeValue(fileWriter, data);
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	
}
