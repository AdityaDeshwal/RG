package Report_Generation.RG;

import com.google.cloud.bigquery.BigQuery;
import com.google.cloud.bigquery.BigQueryOptions;
import com.google.cloud.bigquery.FormatOptions;
import com.google.cloud.bigquery.QueryJobConfiguration;
import com.google.cloud.bigquery.TableDataWriteChannel;
import com.google.cloud.bigquery.TableId;
import com.google.cloud.bigquery.JobId;
import com.google.cloud.bigquery.JobInfo;
import com.google.cloud.bigquery.Job;
import com.google.cloud.bigquery.TableResult;
import com.google.cloud.bigquery.WriteChannelConfiguration;
import com.google.gson.Gson;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.*;

public class SetBTestData {
	private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
	public static Map<String, Object> BTestData=new HashMap<>();
	


    public static String setBTestData(String excelFilePath) {
    	//SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
    	System.out.println("start running");
        try {
        	//DataFormatter dataFormatter = new DataFormatter();

            // Reading data from the Excel file
            FileInputStream file = new FileInputStream(new File(excelFilePath));
            System.out.println("getting sheet 1");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            System.out.println("getting sheet 2");
            Sheet sheet = workbook.getSheet("Upload BTest Info");
            
            //System.out.println("getting sheet");
            
            //Map<String, Object> data = new HashMap<>();
            BTestData.put("test_name", getCellValue(sheet, "B1"));
            BTestData.put("test_code", getCellValue(sheet, "B2"));
            BTestData.put("test_date", getCellValue(sheet, "B3"));
            
            //System.out.println(getCellValue(sheet, "B3"));
            //System.out.println(getCellValue(sheet, 2, 1));
            //data.put("test_name", getCellValue(sheet, "B3"));
            //"C:\Users\adity\Downloads\For Aditya sir.xlsx"
            //System.out.println("getting date");

            // Get the number of subjects
            //System.out.println(getCellValue(sheet,"B4"));
            //Object numsubjects=getCellValue(sheet,"B4");
            int numSubjects =GivingCellValueINT(getCellValue(sheet,"B4"));
           // System.out.println(numSubjects);
            //int numSubjects=3;
            List<Map<String, Object>> subjects = new ArrayList<>();
            int rowNum = 6;
            for (int i = 1; i <= numSubjects; i++) {
                Map<String, Object> subject = new HashMap<>();
                //System.out.println(getCellValue(sheet, rowNum, 1));
                subject.put("subject_name", getCellValue(sheet, rowNum, 1));
                subject.put("position", i);

                int numQTypes = GivingCellValueINT(getCellValue(sheet,rowNum,2));
                //Integer.parseInt(getCellValue(sheet, rowNum, 2).toString());
                List<Map<String, Object>> qTypes = new ArrayList<>();
                for (int j = 1; j <= numQTypes; j++) {
                    Map<String, Object> qType = new HashMap<>();
                    qType.put("q_type_name", getCellValue(sheet, rowNum, 3));
                    qType.put("position", j);
                    qType.put("num_of_qs", getCellValue(sheet, rowNum, 4));
                    qType.put("positive_marks", getCellValue(sheet, rowNum, 5));
                    qType.put("negative_marks", getCellValue(sheet, rowNum, 6));
                    qType.put("has_partial", getCellValue(sheet, rowNum, 7));
                    qType.put("is_best5", getCellValue(sheet, rowNum, 8));
                    qTypes.add(qType);
                    rowNum++;
                    //System.out.println("inner for loop completes");
                }
                subject.put("q_types", qTypes);
                subjects.add(subject);
                //System.out.println("outer for loop completed");
            }
            BTestData.put("subjects", subjects);

            // Get the current timestamp
            Date currentDate = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss z");
            sdf.setTimeZone(TimeZone.getTimeZone("UTC+5:30"));
            BTestData.put("time_stamp", sdf.format(currentDate));
            //System.out.println("function completes");
            // Insert the data into BigQuery
            //System.out.println(BTestData);
//            insertDataIntoBigQuery(bigquery, data);
            file.close();
            return "DONE";
        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }

    private static BigQuery getBigQueryService() {
        return BigQueryOptions.getDefaultInstance().getService();
    }

    private static Object getCellValue(Sheet sheet, String cellReference) {
        int[] cellIndices = getCellIndices(cellReference);
        return getCellValue(sheet, cellIndices[0], cellIndices[1]);
    }

    private static int[] getCellIndices(String cellReference) {
        int row = Integer.parseInt(cellReference.replaceAll("[^0-9]", ""));
        //System.out.println(row);
        int col = cellReference.replaceAll("[^A-Z]", "").charAt(0) - 'A';
        col=col+1;
        //System.out.println(col);
        return new int[]{row, col};
    }
    private static int GivingCellValueINT(Object currvalue) {
    	if(currvalue instanceof Double) {
        	return ((Double) currvalue).intValue();
        }
    	return 0;
    }

    private static Object getCellValue(Sheet sheet, int row, int col) {
    	row=row-1;
    	col=col-1;
        Row r = sheet.getRow(row);
        Cell cell = r.getCell(col);
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    try {
						return dateFormat.format(cell.getDateCellValue());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return null;
        }
    }
}

