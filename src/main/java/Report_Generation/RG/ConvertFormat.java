package Report_Generation.RG;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertFormat{
	private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
	public static Map<String, Object> BTestData=SetBTestData.BTestData;
	public static List<Map<String, Object>> subjects = (List<Map<String, Object>>) BTestData.get("subjects");
	
	public static List<List<Object>> q_info_arr = new ArrayList<>();
    public static List<List<Object>> student_marks_arr_1 = new ArrayList<>();
    public static List<List<Object>> student_info_arr = new ArrayList<>();
    public static List<List<Object>> correction_arr = new ArrayList<>();
	
	public static void ConvertFormat(String excelFilePath) {
		try {
			FileInputStream file = new FileInputStream(new File(excelFilePath));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			Sheet formsheet = workbook.getSheet("Upload Data");
			String inputsheetName=GivingCellValueString(getCellValue(formsheet,"B1"));
			Sheet inputsheet=workbook.getSheet(inputsheetName);
			int q_start_col = letterToColumn(GivingCellValueString(getCellValue(formsheet,"B5")));
			int maxrow=inputsheet.getLastRowNum()+1;
			int maxcol=inputsheet.getRow(0).getLastCellNum();
			System.out.print(maxrow + "\t");
			System.out.print(maxcol + "\t");
			String test_code=GivingCellValueString(getCellValue(formsheet,"B3"));
			
			
			//q_info_arr.add(Arrays.asList("q_id", "test_code", "set_no", "subject", "q_no", "q_setter", "q_type"));
			//student_marks_arr_1.add(Arrays.asList("student_roll_no", "q_id", "marks", "student_answer", "correctness"));
			//student_info_arr.add(Arrays.asList("roll_no", "name"));
			
		    
		    Object[][] info_range = new Object[maxrow][maxcol];
	        for (int i = 0; i < maxrow; i++) {
	            Row row = inputsheet.getRow(i);
	            for (int j = 0; j < maxcol; j++) {
	                Cell cell = row.getCell(j);
	                if (cell != null) {
	                    switch (cell.getCellType()) {
	                        case STRING:
	                            info_range[i][j] = cell.getStringCellValue();
	                            break;
	                        case NUMERIC:
	                            if (DateUtil.isCellDateFormatted(cell)) {
	                                info_range[i][j] = cell.getDateCellValue();
	                            } else {
	                                info_range[i][j] = cell.getNumericCellValue();
	                            }
	                            break;
	                        case BOOLEAN:
	                            info_range[i][j] = cell.getBooleanCellValue();
	                            break;
	                        case FORMULA:
	                            info_range[i][j] = cell.getCellFormula();
	                            break;
	                        case BLANK:
	                            info_range[i][j] = "";
	                            break;
	                        default:
	                            info_range[i][j] = "Unsupported Cell Type";
	                    }
	                } else {
	                    info_range[i][j] = "";
	                }
	            }
	        }
	        
//	        for (Object[] row : info_range) {
//	            for (Object cellValue : row) {
//	                System.out.print(cellValue + "\t");
//	            }
//	            System.out.println();
//	        }
		    
		    
		    
		    int set_col = letterToColumn(GivingCellValueString(getCellValue(formsheet,"B4"))) - 1;
		    Map<String, List<Integer>> best5qs = new HashMap<>();
		    int col = q_start_col - 1;
		    List<String> possible_set_nums = new ArrayList<>();
		    possible_set_nums = Arrays.asList("A", "B", "C");
		   // String medium = PropertiesService.getScriptProperties().getProperty("medium");
//		    if ("Offline".equals(medium)) {
//		        possible_set_nums = Arrays.asList("A", "B", "C");
//		    } else if ("Online".equals(medium)) {
//		        possible_set_nums = Arrays.asList("O");
//		    }
		    System.out.println(col);
		    for (Map<String, Object> subj : subjects) {
		        String subject_name = (String) subj.get("subject_name");
		        int qs_till_now = 0;
		        List<Map<String, Object>> q_types = (List<Map<String, Object>>) subj.get("q_types");
		        for (Map<String, Object> q_type : q_types) {
		            String q_type_name = (String) q_type.get("q_type_name");
		            boolean isbest5 = "true".equals(q_type.get("is_best5"));
		            int best5start = 1 + qs_till_now;
		            //System.out.println("error in 1st only");
		            for (int q_no = 1 + qs_till_now; q_no <= (int) ((Double) q_type.get("num_of_qs") + qs_till_now); q_no++) {
		                for (String set_no : possible_set_nums) {
		                    String q_id = createQId(test_code, set_no, subject_name, q_no);
		                    q_info_arr.add(Arrays.asList(q_id, test_code, set_no, subject_name, q_no, "", q_type_name));
		                }

		                for (int row = 2; row < maxrow; row++) {/*leaving top 2 rows as they contain the name of the columns and not information*/
		                    String roll_no = info_range[row][0].toString();
		                    String set_no = info_range[row][set_col].toString();

		                    if (col == q_start_col - 1) {
		                        if (!possible_set_nums.contains(set_no)) {
		                            correction_arr.add(Arrays.asList(roll_no));
		                            continue;
		                        }
		                    }

		                    String q_id = createQId(test_code, set_no, subject_name, q_no);
		                    Object marks = info_range[row][col + 2];
		                    String correctness = "NOT ANSWERED";
		                    if (info_range[row][col + 3].equals(1.0)) correctness = "CORRECT";
		                    else if (info_range[row][col + 4].equals(1.0)) correctness = "NOT CORRECT";
		                    else if (info_range[row][col + 5].equals(1.0)) correctness = "PARTIALLY CORRECT";

		                    if (isbest5) {
		                        if (q_no == best5start) {
		                            best5qs.put(roll_no, new ArrayList<>());
		                        }
		                        if (best5qs.get(roll_no).size() >= 5) {
		                            marks = 0;
		                            correctness = "NOT ANSWERED";
		                        } else if (!"NOT ANSWERED".equals(correctness)) {
		                            best5qs.get(roll_no).add(q_no);
		                        }
		                    }

		                    student_marks_arr_1.add(Arrays.asList(roll_no, q_id, marks, "", correctness));
		                }
		                col += 7;
		                //System.out.println(col);
		            }
		            //System.out.println("error in 2nd only");
		            qs_till_now +=(int) ((Double) q_type.get("num_of_qs") + 0);
		        }
		    }
		    

	        for (int row = 2; row < maxrow; row++) {
	            Object roll_no = info_range[row][0];
	            Object name = info_range[row][1];
	            List<Object> student_info = new ArrayList<>();
	            student_info.add(roll_no);
	            student_info.add(name);
	            student_info_arr.add(student_info);
	        }
	        
//	        for (List<Object> student : student_info_arr) {
//	            for (Object info : student) {
//	                System.out.print(info + " ");
//	            }
//	            System.out.println(); // Move to the next line after each student's info
//	        }
	        
//	        for (List<Object> student : student_marks_arr_1) {
//	            for (Object data : student) {
//	                System.out.print(data + "\t");
//	            }
//	            System.out.println();
//	        }
	        
//	        for (List<Object> question : q_info_arr) {
//	            for (Object info : question) {
//	                System.out.print(info + " ");
//	            }
//	            System.out.println(); // Move to the next line after each question's info
//	        }
		    
//	        PopulateSheet.populateSheet(workbook, "Student_Marks",student_marks_arr_1 , excelFilePath, test_code);
//	        PopulateSheet.populateSheet(workbook, "Student_Info",student_info_arr , excelFilePath, test_code);
//	        PopulateSheet.populateSheet(workbook, "Question_Info",q_info_arr , excelFilePath, test_code);
	        
	        file.close();
	        workbook.close();
	        System.gc();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
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
    private static String GivingCellValueString(Object currvalue) {
    	if (currvalue instanceof String) {
            return (String) currvalue;
        }
    	return "";
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
    
    public static int letterToColumn(String letter) {
        int column = 0;
        int length = letter.length();
        for (int i = 0; i < length; i++) {
            column += (letter.charAt(i) - 'A' + 1) * Math.pow(26, length - i - 1);
        }
        return column;
    }

    // Convert a number to a two-digit string
    public static String twoDigitNumber(int num) {
        if (num < 10) {
            return "0" + num;
        } else if (num < 100) {
            return Integer.toString(num);
        } else {
            return "XX";  // In case the number has more than 2 digits
        }
    }
    
    public static String createQId(String testCode, String setNo, String subjCode, int qNo) {
        return testCode + setNo + subjCode + twoDigitNumber(qNo);
    }
}
