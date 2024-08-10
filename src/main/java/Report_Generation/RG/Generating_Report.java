package Report_Generation.RG;

import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberTickUnit;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.StackedBarRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.title.TextTitle;
import org.jfree.chart.ui.RectangleEdge;
import org.jfree.chart.ui.RectangleInsets;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.ui.RefineryUtilities;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.WorkbookDocumentImpl;

import com.fasterxml.jackson.databind.ser.std.StdKeySerializers.Default;
import com.graphbuilder.math.func.LnFunction;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Section;
import com.itextpdf.text.pdf.DeviceNColor;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chapter;

import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.NumberAxis;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.stream.Collectors;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;

import org.apache.commons.math3.analysis.function.Max;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xddf.usermodel.text.TextAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;
import org.apache.xmlbeans.impl.common.XmlStreamUtils;
import org.checkerframework.checker.units.qual.min;






public class Generating_Report {
	private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
	private static final SimpleDateFormat IndiandateFormat = new SimpleDateFormat("dd-MM-yyyy");
	public static Map<String, Object> BTest_data = SetBTestData.BTestData;
	private static final Map<String, String> subjFullForm = Map.of(
	        "PHY", "Physics",
	        "CHEM", "Chemistry",
	        "MATH", "Mathematics",
	        "LOG", "Logic",
	        "COD", "Coding"
	    );
	private static Map<String, Double> percentiles=new HashMap<>();
	private static Map<String, Double> averages=new HashMap<>();
	private static Map<String, Double> neg_averages=new HashMap<>();
	private static List<Map<String, Object>> finalOutput=new ArrayList<Map<String,Object>>();
	private static Map<String, Map<String, Map<String, Map<String, Double>>>> ques_analysis = new HashMap<>();
	private static Map<String, Map<String, Object>> t2=new HashMap<String, Map<String,Object>>();
	private static DecimalFormat df = new DecimalFormat("#.##");
	private static Map<String, XSSFChart> chartMap = new HashMap<>();
	private static boolean islastSet=false;
	private static String excelPath="";
	private static String BTest_Name=(String) BTest_data.get("test_name");;
	private static String folderFilePath="";
	private static ProcessBuilder processBuilder;
	private static Double total_marks_per_80=0.0;
//    for (Double value : percentiles.values()) {
//        total_marks_per_80+= value;
//    }
	private static String test_date_unparsed=(String)(BTest_data.get("test_date"));
	
	  private static String convertDateFormat(String dateString) {
	        try {
	            Date date = dateFormat.parse(dateString);
	            return IndiandateFormat.format(date);
	        } catch (ParseException e) {
	            return "Invalid Date";
	        }
	    }
	  private static String test_date=convertDateFormat(test_date_unparsed);
	
	public static void adjusting_data(String excelFilePath) { 
		try {
			FileInputStream file = new FileInputStream(new File(excelFilePath));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			Sheet formsheet = workbook.getSheet("Create Report"); 
			String testCode = GivingCellValueString(getCellValue(formsheet,"B1"));
			String medium = GivingCellValueString(getCellValue(formsheet,"B2"));
			String equality = (medium.equals("Online")) ? "=" : "!=";
	        //excelPath=excelFilePath;
			
			System.out.println("Good Before Strp1");
	        // Step 1: Join and filter
	        List<Map<String, Object>> t1 = ConvertFormat.student_marks_arr_1.stream()
	                .map(studentMark -> {
	                    Map<String, Object> joinedData = new HashMap<>();
	                    joinedData.put("student_roll_no", studentMark.get(0));
	                    //joinedData.put("q_id", studentMark.get(1));
	                    joinedData.put("marks", studentMark.get(2));
	                    joinedData.put("correctness", studentMark.get(4));

	                    // Join with q_info_arr
	                    ConvertFormat.q_info_arr.stream()
	                            .filter(q -> q.get(0).equals(studentMark.get(1)))
	                            .findFirst()
	                            .ifPresent(q -> {
	                                joinedData.put("subject", q.get(3));
	                                joinedData.put("q_type", q.get(6));
	                                joinedData.put("set_no", q.get(2));
	                                joinedData.put("q_no", q.get(4));
//	                                joinedData.put("is_best5", q.get(7));
//	                                System.out.println(q.get(7));
	                            });

	                    // Join with student_info_arr
	                    ConvertFormat.student_info_arr.stream()
	                            .filter(s -> s.get(0).equals(studentMark.get(0)))
	                            .findFirst()
	                            .ifPresent(s -> joinedData.put("name", s.get(1)));

	                    return joinedData;
	                })
	                .filter(e -> {
	                    String setNo = e.get("set_no").toString();
	                    return equality.equals("=") ? setNo.equals("O") : !setNo.equals("O");
	                })
	                .collect(Collectors.toList());
//	        e.get("test_code").equals(testCode) &&
//	        .filter(e -> e.get("set_no").toString().equals(equality))
	        
	        //t1.forEach(System.out::println);
	        ConvertFormat.student_marks_arr_1.clear();
			ConvertFormat.student_info_arr.clear();
			ConvertFormat.q_info_arr.clear();
			ConvertFormat.correction_arr.clear();
			System.gc();
	        System.out.println("Step 1 working fine");
	        // Step 2: Pivot and aggregate
	        t2 = t1.stream()
	        	    .collect(Collectors.groupingBy(
	        	        e -> e.get("student_roll_no").toString(),
	        	        Collectors.collectingAndThen(
	        	            Collectors.toList(),
	        	            list -> {
	        	                Map<String, Object> aggregatedData = new HashMap<>();
	        	                aggregatedData.put("student_roll_no", list.get(0).get("student_roll_no"));
	        	                aggregatedData.put("name", list.get(0).get("name"));
	        	                aggregatedData.put("set_no", list.get(0).get("set_no"));

	        	                // Initialize aggregated counts and totals for each subject
	        	                Map<String, Double> totalMarksBySubject = new HashMap<>();
	        	                Map<String, Double> totalPositiveMarksBySubject = new HashMap<>();
	        	                Map<String, Double> totalNegativeMarksBySubject = new HashMap<>();
	        	                Map<String, Integer> totalNumQsBySubject = new HashMap<>();
	        	                Map<String, Integer> totalNumAttemptedBySubject = new HashMap<>();
	        	                Map<String, Integer> totalNumCorrectBySubject = new HashMap<>();
	        	                Map<String, Integer> totalNumIncorrectBySubject = new HashMap<>();

	        	                // Aggregate data across all subjects for the student
	        	                for (Map<String, Object> entry : list) {
	        	                    String subject = entry.get("subject").toString();
	        	                    double marks = (Double) entry.get("marks");

	        	                    // Initialize subject-specific maps if not already present
	        	                    totalMarksBySubject.putIfAbsent(subject, 0.0);
	        	                    totalPositiveMarksBySubject.putIfAbsent(subject, 0.0);
	        	                    totalNegativeMarksBySubject.putIfAbsent(subject, 0.0);
	        	                    totalNumQsBySubject.putIfAbsent(subject, 0);
	        	                    totalNumAttemptedBySubject.putIfAbsent(subject, 0);
	        	                    totalNumCorrectBySubject.putIfAbsent(subject, 0);
	        	                    totalNumIncorrectBySubject.putIfAbsent(subject, 0);

	        	                    // Aggregate marks and counts for the specific subject
	        	                    totalMarksBySubject.put(subject, totalMarksBySubject.get(subject) + marks);
	        	                    totalPositiveMarksBySubject.put(subject, totalPositiveMarksBySubject.get(subject) + Math.max(0, marks));
	        	                    totalNegativeMarksBySubject.put(subject, totalNegativeMarksBySubject.get(subject) + Math.min(0, marks));
	        	                    totalNumQsBySubject.put(subject, totalNumQsBySubject.get(subject) + 1);
	        	                    if (!entry.get("correctness").equals("NOT ANSWERED")) {
	        	                        totalNumAttemptedBySubject.put(subject, totalNumAttemptedBySubject.get(subject) + 1);
	        	                    }
	        	                    if (entry.get("correctness").equals("CORRECT")) {
	        	                        totalNumCorrectBySubject.put(subject, totalNumCorrectBySubject.get(subject) + 1);
	        	                    } else if (entry.get("correctness").equals("NOT CORRECT")) {
	        	                        totalNumIncorrectBySubject.put(subject, totalNumIncorrectBySubject.get(subject) + 1);
	        	                    }
	        	                }

	        	                // Put aggregated totals for each subject into aggregatedData map
	        	                for (String subject : totalMarksBySubject.keySet()) {
	        	                    aggregatedData.put(subject + "_total_marks", totalMarksBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_positive_marks", totalPositiveMarksBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_negative_marks", totalNegativeMarksBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_num_qs", totalNumQsBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_num_attempted", totalNumAttemptedBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_num_correct", totalNumCorrectBySubject.get(subject));
	        	                    aggregatedData.put(subject + "_num_incorrect", totalNumIncorrectBySubject.get(subject));
	        	                }

	        	                return aggregatedData;
	        	            }
	        	        )
	        	    ));;


	         //Print t2 to verify the results
//	        t2.forEach((key, data) -> {
//	            System.out.println("Key: " + key);
//	            data.forEach((innerKey, value) -> System.out.println("  " + innerKey + ": " + value));
//	        });     
//	        
	        System.gc();
	        System.out.println("Step 2 working fine");
//	        // Step 3: Further aggregate by student and set number
	        
	        List<Map<String, Object>> subjectsList = (List<Map<String, Object>>) BTest_data.get("subjects");

	        // Map to hold subject averages and percentiles
//	        Map<String, Double> averages = new HashMap<>();
//	        Map<String, Double> percentiles = new HashMap<>();
//	        neg_averages=new HashMap<>();
	        // Calculate averages and 80th percentiles for each subject
	        for (Map<String, Object> subjectMap : subjectsList) {
	            String subjectName = (String) subjectMap.get("subject_name");
	            averages.putAll(calculateAverages(t2, subjectName));
	            neg_averages.putAll(calculateNegativeAverages(t2, subjectName));
	            percentiles.putAll(calculatePercentiles(t2, subjectName, 0.80));
	        }

	        // Create t3
	        System.gc();
	        List<Map<String, Object>> t3 = t2.values().stream()
	            .map(aggregatedData -> {
	                Map<String, Object> rowData = new HashMap<>();
	                rowData.put("student_roll_no", aggregatedData.get("student_roll_no"));
	                rowData.put("name", aggregatedData.get("name"));
	                rowData.put("set_no", aggregatedData.get("set_no"));

	                for (Map<String, Object> subjectMap : subjectsList) {
	                    String subjectName = (String) subjectMap.get("subject_name");
	                    double totalMarks = (Double) aggregatedData.getOrDefault(subjectName + "_total_marks", 0.0);
	                    double avgMarks = averages.get("avg_" + subjectName);
	                    double percentile80 = percentiles.get(subjectName + "_80th_percentile");

	                    rowData.put(subjectName + "_total_marks", totalMarks);
	                    rowData.put(subjectName + "_marks_per_avg", totalMarks / avgMarks);
	                    rowData.put(subjectName + "_marks_per_80", totalMarks / percentile80);
	                    rowData.put(subjectName + "_positive_marks", aggregatedData.getOrDefault(subjectName + "_positive_marks", 0.0));
	                    rowData.put(subjectName + "_negative_marks", aggregatedData.getOrDefault(subjectName + "_negative_marks", 0.0));
	                    rowData.put(subjectName + "_num_qs", aggregatedData.getOrDefault(subjectName + "_num_qs", 0));
	                    rowData.put(subjectName + "_num_attempted", aggregatedData.getOrDefault(subjectName + "_num_attempted", 0));
	                    rowData.put(subjectName + "_num_correct", aggregatedData.getOrDefault(subjectName + "_num_correct", 0));
	                }

	                return rowData;
	            })
	            .collect(Collectors.toList());
	        //t3.forEach(System.out::println);
	    
	        System.out.println("Step 3 working fine");
//	        // Step 4: Calculate percentages
	        t2.clear();
	        System.gc();
	        //System.out.println(t1);
	        Map<String, Map<String, Map<String, Map<String, Object>>>> groupedData = t1.stream()
	                .collect(Collectors.groupingBy(
	                    e -> e.get("student_roll_no").toString(),
	                    Collectors.groupingBy(
	                        e -> e.get("subject").toString(),
	                        Collectors.groupingBy(
	                            e -> e.get("q_type").toString(),
	                            Collectors.collectingAndThen(
	                                Collectors.toList(),
	                                list -> {
	                                    int totalQs = list.size();
//	                                    boolean isBest5 = list.stream().anyMatch(entry -> Boolean.TRUE.equals(entry.get("is_best5")));
//	                                    //System.out.println(isBest5);
//	                                    if(isBest5)totalQs=5;
	                                    long correctCount = list.stream().filter(entry -> entry.get("correctness").equals("CORRECT")).count();
	                                    long incorrectCount = list.stream().filter(entry -> entry.get("correctness").equals("NOT CORRECT")).count();
	                                    long attemptedCount = list.stream().filter(entry -> !entry.get("correctness").equals("NOT ANSWERED")).count();

	                                    double correctPerc = totalQs > 0 ? (correctCount * 100.0 / totalQs) : 0.0;
	                                    double incorrectPerc = totalQs > 0 ? (incorrectCount * 100.0 / totalQs) : 0.0;
	                                    double attemptedPerc = totalQs > 0 ? (attemptedCount * 100.0 / totalQs) : 0.0;

	                                    Map<String, Object> result = new HashMap<>();
	                                    result.put("correct_perc", correctPerc);
	                                    result.put("incorrect_perc", incorrectPerc);
	                                    result.put("attempted_perc", attemptedPerc);
	                                    return result;
	                                }
	                            )
	                        )
	                    )
	                ));

	            // Create t4
	            List<Map<String, Object>> t4 = groupedData.entrySet().stream()
	                .map(studentEntry -> {
	                    String studentRollNo = studentEntry.getKey();
	                    Map<String, Map<String, Map<String, Object>>> subjectsMap = studentEntry.getValue();

	                    Map<String, Object> rowData = new HashMap<>();
	                    rowData.put("student_roll_no", studentRollNo);
	                    rowData.put("name", t1.stream().filter(e -> e.get("student_roll_no").toString().equals(studentRollNo)).findFirst().orElse(new HashMap<>()).get("name"));
	                    rowData.put("set_no", t1.stream().filter(e -> e.get("student_roll_no").toString().equals(studentRollNo)).findFirst().orElse(new HashMap<>()).get("set_no"));

	                    subjectsList.forEach(subject -> {
	                        String subjectName = subject.get("subject_name").toString();
	                        List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subject.get("q_types");

	                        qTypes.forEach(qType -> {
	                            String qTypeName = qType.get("q_type_name").toString();

	                            Map<String, Object> percData = subjectsMap.getOrDefault(subjectName, Collections.emptyMap())
	                                .getOrDefault(qTypeName, new HashMap<>());

	                            rowData.put(subjectName + "_" + qTypeName + "_correct_perc", percData.getOrDefault("correct_perc", 0.0));
	                            rowData.put(subjectName + "_" + qTypeName + "_incorrect_perc", percData.getOrDefault("incorrect_perc", 0.0));
	                            rowData.put(subjectName + "_" + qTypeName + "_attempted_perc", percData.getOrDefault("attempted_perc", 0.0));
	                        });
	                    });

	                    return rowData;
	                })
	                .collect(Collectors.toList());
	        
	       // t4.forEach(System.out::println);
////	        
	        System.out.println("Step 4 working fine");
	        groupedData.clear();
	        System.gc();
//
//	        // Step 5: Final combination and sorting
	        Map<String, Map<String, Map<String, Map<String, Object>>>> groupedDatafort5 = t1.stream()
	                .collect(Collectors.groupingBy(
	                    e -> e.get("student_roll_no").toString(),
	                    Collectors.groupingBy(
	                        e -> e.get("subject").toString(),
	                        Collectors.groupingBy(
	                            e -> e.get("q_type").toString(),
	                            Collectors.collectingAndThen(
	                                Collectors.toList(),
	                                list -> {
	                                    long correctCount = list.stream().filter(entry -> entry.get("correctness").equals("CORRECT")).count();
	                                    long incorrectCount = list.stream().filter(entry -> entry.get("correctness").equals("NOT CORRECT")).count();
	                                    long notAnsweredCount = list.stream().filter(entry -> entry.get("correctness").equals("NOT ANSWERED")).count();

	                                    Map<String, Object> result = new HashMap<>();
	                                    result.put("correct", correctCount);
	                                    result.put("incorrect", incorrectCount);
	                                    result.put("not_answered", notAnsweredCount);
	                                    return result;
	                                }
	                            )
	                        )
	                    )
	                ));
	        
	        Map<String, Map<String, Map<String, String>>> ques_analysis_per_student = t1.stream()
	                .collect(Collectors.groupingBy(
	                        e -> e.get("student_roll_no").toString(),
	                        Collectors.toMap(
	                                e -> e.get("subject").toString(),
	                                e -> {
	                                    String statusKey = "status_" + e.get("subject").toString() + "_" + e.get("q_no").toString();
	                                    String statusValue = e.get("correctness").toString();
	                                    Map<String, String> subjectStatus = new HashMap<>();
	                                    subjectStatus.put(statusKey, statusValue);
	                                    return subjectStatus;
	                                },
	                                (existing, replacement) -> {
	                                    existing.putAll(replacement);
	                                    return existing;
	                                }
	                        )
	                ));

	            // Create t5
	            List<Map<String, Object>> t5 = groupedDatafort5.entrySet().stream()
	                .map(studentEntry -> {
	                    String studentRollNo = studentEntry.getKey();
	                    Map<String, Map<String, Map<String, Object>>> subjectsMap = studentEntry.getValue();

	                    Map<String, Object> rowData = new HashMap<>();
	                    rowData.put("student_roll_no", studentRollNo);
	                    rowData.put("name", t1.stream().filter(e -> e.get("student_roll_no").toString().equals(studentRollNo)).findFirst().orElse(new HashMap<>()).get("name"));
	                    rowData.put("set_no", t1.stream().filter(e -> e.get("student_roll_no").toString().equals(studentRollNo)).findFirst().orElse(new HashMap<>()).get("set_no"));

	                    subjectsList.forEach(subject -> {
	                        String subjectName = subject.get("subject_name").toString();
	                        List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subject.get("q_types");

	                        long subjectCorrect = 0;
	                        long subjectIncorrect = 0;
	                        long subjectNotAnswered = 0;

	                        for (Map<String, Object> qType : qTypes) {
	                            String qTypeName = qType.get("q_type_name").toString();

	                            Map<String, Object> correctnessData = subjectsMap.getOrDefault(subjectName, Collections.emptyMap())
	                                .getOrDefault(qTypeName, new HashMap<>());

	                            long correct = (long) correctnessData.getOrDefault("correct", 0L);
	                            long incorrect = (long) correctnessData.getOrDefault("incorrect", 0L);
	                            long notAnswered = (long) correctnessData.getOrDefault("not_answered", 0L);

	                            subjectCorrect += correct;
	                            subjectIncorrect += incorrect;
	                            subjectNotAnswered += notAnswered;

	                            rowData.put(subjectName + "_" + qTypeName + "_correctness", "Correct: " + correct + ", Incorrect: " + incorrect + ", Not Answered: " + notAnswered);
	                        }

	                        rowData.put(subjectName + "_correctness", "Correct: " + subjectCorrect + ", Incorrect: " + subjectIncorrect + ", Not Answered: " + subjectNotAnswered);
	                    });

	                    return rowData;
	                })
	                .collect(Collectors.toList());

	            // Print t5 for verification
	           // t5.forEach(System.out::println);
	            t1.clear();
	            groupedDatafort5.clear();
	            System.gc();
	            
	            List<String> subj_names = subjectsList.stream()
	                    .map(sub -> sub.get("subject_name").toString())
	                    .collect(Collectors.toList());

	                List<String> q_type_names = subjectsList.stream()
	                    .flatMap(sub -> ((List<Map<String, Object>>) sub.get("q_types")).stream().map(qt -> qt.get("q_type_name").toString()))
	                    .distinct()
	                    .collect(Collectors.toList());

	                // Join t3, t4, and t5 based on set_no, student_roll_no, and name
	                finalOutput= t3.stream()
	                    .map(t3Row -> {
	                        String setNo = t3Row.get("set_no").toString();
	                        String studentRollNo = t3Row.get("student_roll_no").toString();
	                        String name = t3Row.get("name").toString();

	                        // Find corresponding rows in t4 and t5
	                        Map<String, Object> t4Row = t4.stream()
	                            .filter(row -> row.get("set_no").toString().equals(setNo)
	                                    && row.get("student_roll_no").toString().equals(studentRollNo)
	                                    && row.get("name").toString().equals(name))
	                            .findFirst()
	                            .orElse(new HashMap<>());

	                        Map<String, Object> t5Row = t5.stream()
	                            .filter(row -> row.get("set_no").toString().equals(setNo)
	                                    && row.get("student_roll_no").toString().equals(studentRollNo)
	                                    && row.get("name").toString().equals(name))
	                            .findFirst()
	                            .orElse(new HashMap<>());

	                        Map<String, Object> finalRow = new HashMap<>(t3Row);

	                        // Add fields from t4
	                        for (String subj : subj_names) {
	                            for (String qType : q_type_names) {
	                                finalRow.put(subj + "_" + qType + "_correct_perc", t4Row.getOrDefault(subj + "_" + qType + "_correct_perc", 0));
	                                finalRow.put(subj + "_" + qType + "_incorrect_perc", t4Row.getOrDefault(subj + "_" + qType + "_incorrect_perc", 0));
	                                finalRow.put(subj + "_" + qType + "_attempted_perc", t4Row.getOrDefault(subj + "_" + qType + "_attempted_perc", 0));
	                            }
	                        }

	                        // Add fields from t5
	                        for (String subj : subj_names) {
	                            finalRow.put(subj + "_correctness", t5Row.getOrDefault(subj + "_correctness", ""));
	                            for (String qType : q_type_names) {
	                                finalRow.put(subj + "_" + qType + "_correctness", t5Row.getOrDefault(subj + "_" + qType + "_correctness", ""));
	                            }
	                        }

	                        // Add question-wise status fields from t5
	                        
	                        if (ques_analysis_per_student.containsKey(studentRollNo)) {
	                            Map<String, Map<String, String>> studentStatus = ques_analysis_per_student.get(studentRollNo);
	                            for (String subj : studentStatus.keySet()) {
	                                for (String statusKey : studentStatus.get(subj).keySet()) {
	                                    finalRow.put(statusKey, studentStatus.get(subj).get(statusKey));
	                                }
	                            }
	                        }


	                        // Calculate additional fields
	                        for (String subj : subj_names) {
	                            double totalMarks = (double) finalRow.getOrDefault(subj + "_total_marks", 0.0);
	                            double avgMarks = averages.get("avg_" + subj);
	                            		//t3.stream().mapToDouble(row -> (double) row.getOrDefault(subj + "_total_marks", 0.0)).average().orElse(0.0);
	                            double percentile80 = percentiles.get(subj + "_80th_percentile");
	                            		//t3.stream().mapToDouble(row -> (double) row.getOrDefault(subj + "_total_marks", 0.0)).sorted().skip((long) (0.8 * t3.size())).findFirst().orElse(0.0);

	                            finalRow.put("marks_per_avg_" + subj, totalMarks / (avgMarks == 0 ? 1 : avgMarks));
	                            finalRow.put("marks_per_80_" + subj, totalMarks / (percentile80 == 0 ? 1 : percentile80));
	                        }

	                        return finalRow;
	                    })
	                    .collect(Collectors.toList());
	                
	                t1.clear();
	                t2.clear();
	                t3.clear();
	                t4.clear();
	                t5.clear();
	                System.gc();

	                // Process finalOutput
	                for (Map<String, Object> studentData : finalOutput) {
	                    String setNo = studentData.get("set_no").toString();
	                    String studentRollNo = studentData.get("student_roll_no").toString();

	                    for (Map.Entry<String, Object> entry : studentData.entrySet()) {
	                        String key = entry.getKey();
	                        if (key.startsWith("status_")) {
	                            String[] parts = key.split("_");
	                            if (parts.length >= 3) {
	                                String subject = parts[1];
	                                String question = parts[1] + "_" + parts[2];
	                                String status = entry.getValue().toString().trim();

	                                // Ensure subject, setNo, and question are initialized
	                                ques_analysis.computeIfAbsent(subject, k -> new HashMap<>())
	                                        .computeIfAbsent(setNo, k -> new HashMap<>())
	                                        .computeIfAbsent(question, k -> initializeStats());

	                                // Update statistics
	                                Map<String, Double> stats = ques_analysis.get(subject).get(setNo).get(question);
	                                stats.merge("total", 1.0, Double::sum);

	                                if (!status.equals("NOT ANSWERED")) {
	                                    stats.merge("attempted", 1.0, Double::sum);
	                                    if (status.equals("CORRECT")) {
	                                        stats.merge("correct", 1.0, Double::sum);
	                                    }
	                                }
	                            } else {
	                                System.out.println("Invalid key format: " + key);
	                            }
	                        }
	                    }
	                }

	                // Calculate percentages
	                for (String subject : ques_analysis.keySet()) {
	                    for (String setNo : ques_analysis.get(subject).keySet()) {
	                        for (String question : ques_analysis.get(subject).get(setNo).keySet()) {
	                            Map<String, Double> stats = ques_analysis.get(subject).get(setNo).get(question);
	                            double total = stats.getOrDefault("total", 0.0);
	                            double attempted = stats.getOrDefault("attempted", 0.0);
	                            double correct = stats.getOrDefault("correct", 0.0);

	                            if (total > 0) {
	                                stats.put("attempted_perc", (attempted / total) * 100);
	                                stats.put("correct_perc", (correct / total) * 100);
	                            } else {
	                                stats.put("attempted_perc", 0.0);
	                                stats.put("correct_perc", 0.0);
	                            }
	                        }
	                    }
	                }

	                // Print or use the ques_analysis map with percentages
//	                System.out.println("Ques Analysis with Sets:");
//	                System.out.println(ques_analysis);

	                for (String subject : ques_analysis.keySet()) {
	                    for (String setNo : ques_analysis.get(subject).keySet()) {
	                        Map<String, Map<String, Double>> questions = ques_analysis.get(subject).get(setNo);

	                        // Convert the questions map to a list for sorting
	                        List<Map.Entry<String, Map<String, Double>>> questionList = new ArrayList<>(questions.entrySet());

	                        // Sort the questionList based on attempted_perc (descending order)
	                        questionList.sort((q1, q2) -> {
	                            double perc1 = q1.getValue().getOrDefault("correct_perc", 0.0);
	                            double perc2 = q2.getValue().getOrDefault("correct_perc", 0.0);
	                            return Double.compare(perc2, perc1); // Descending order
	                        });

	                        // Reconstruct the sorted map
	                        LinkedHashMap<String, Map<String, Double>> sortedQuestions = new LinkedHashMap<>();
	                        for (Map.Entry<String, Map<String, Double>> entry : questionList) {
	                            sortedQuestions.put(entry.getKey(), entry.getValue());
	                        }

	                        // Replace the unsorted questions map with the sorted one
	                        ques_analysis.get(subject).put(setNo, sortedQuestions);
	                    }
	                }
	                workbook.close();
//	                
	                System.gc();
	                createFolder();	               
	           
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void createFolder() {
		String folderName = BTest_Name;
		 String desktopPath = getDesktopPath();
	        if (desktopPath == null) {
	            throw new IllegalStateException("Desktop path could not be found");
	        }

	        // Create the new folder path
	        Path newFolderPath = Paths.get(desktopPath, folderName);

	        try {
	            // Create the folder if it doesn't exist
	            if (!Files.exists(newFolderPath)) {
	                Files.createDirectory(newFolderPath);
	            }
	            System.out.println("Folder created at: " + newFolderPath.toString());
	            folderFilePath = newFolderPath.toString();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	}
	
	public static void createReportsPdf() {
		int total=finalOutput.size();
		for(int i=0;i<total;i++) {
			long stTime = System.currentTimeMillis();
        	directPdf(i);
        	App.updateProgress(i+1 + "/" + total);
        	System.out.println(i+1 +"-" +(System.currentTimeMillis()-stTime));
        }
	}
	public static List<String> quotes=new ArrayList<>();
	static {
		quotes.add("“A person who never made a mistake never tried anything new.” —Albert Einstein");
		quotes.add("“Discipline will take you places motivation can't.” –Anonymus");
		quotes.add("“Procrastination makes easy things hard and hard things harder.” —Mason Cooley");
		quotes.add("“The best way to predict your future is to create it.” —Abraham Lincoln");
		quotes.add("“He who asks a question is a fool for five minutes; he who does not ask a question remains a fool forever.” —Chinese Proverb");
		quotes.add("“The biggest room in the world is the room for improvement.” -Helmut Schmidt");
		quotes.add("“End is not the end if fact E.N.D. Means “Efforts Never Dies. – Dr. A.P.J. Abdul Kalam");
		quotes.add("“If you get tired, learn to rest not to quit.” – Banksy");
	}
	
	public static void directPdf(Integer ind) { 
		Map<String, Object> currdata=finalOutput.get(ind);

		String pdfFilePath=folderFilePath + "\\" + currdata.get("student_roll_no") + "_" + BTest_Name + ".pdf";
		
		Random random = new Random();
		List<String> currQuotes=new ArrayList<>(quotes);
		try {
            // Initialize PDF writer and document
            Document document = new Document();
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(pdfFilePath));
        
            
           // String footerText1 = "Bakliwal Tutorials ( Where BesT teachers meet BesT students )";
            
            final com.itextpdf.text.Font footerFont = FontFactory.getFont(FontFactory.TIMES_ITALIC, 8,BaseColor.BLACK);

            // Register the page event for footer and page numbers
            writer.setPageEvent(new PdfPageEventHelper() {
                @Override
                public void onEndPage(PdfWriter writer, Document document) {
                    PdfPTable footer = new PdfPTable(1);
                    try {
                    	int randomIndex = random.nextInt(currQuotes.size());
                        String footerText1 = currQuotes.get(randomIndex);
                        currQuotes.remove(randomIndex);
                        // Set the total width of the table to match the page width
                        footer.setTotalWidth(document.getPageSize().getWidth() - document.leftMargin() - document.rightMargin());
                        footer.setLockedWidth(true);

                        // Add the footer text and page number
                        PdfPCell cell = new PdfPCell(new Paragraph(writer.getPageNumber() + "\n\n" + footerText1 , footerFont));
                        cell.setBorder(Rectangle.NO_BORDER);
                        cell.setBorderWidthTop(1f);
                        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cell.setVerticalAlignment(Element.ALIGN_BOTTOM);
                        footer.addCell(cell);

                        // Write the footer to the document
                        footer.writeSelectedRows(0, -1, document.leftMargin(), document.bottomMargin(), writer.getDirectContent());
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            });
            
            document.open();
            
            Chapter chapter_mark_analysis = new Chapter(1);
            chapter_mark_analysis.setNumberDepth(0);
           
            InputStream logoStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Bakliwal_Logo.png");
            byte[] logoBytes = IOUtils.toByteArray(logoStream);
            Image logoImage = Image.getInstance(logoBytes);
            logoImage.setAbsolutePosition(20, document.top()-20); // Positioning the image on the page
            logoImage.scaleToFit(45, 45);
            //chapter_mark_analysis.add(logoImage);
            
            float[] columnWidths_mainHeadingTable = {15f,70f,15f};
            PdfPTable main_heading_table = new PdfPTable(columnWidths_mainHeadingTable);
            main_heading_table.setWidthPercentage(100); 
            
            PdfPCell logoCell=new PdfPCell();
            logoCell.addElement(logoImage);
            logoCell.setBorderColor(BaseColor.BLACK);
            logoCell.setBackgroundColor(BaseColor.BLACK);
            main_heading_table.addCell(logoCell);
            
            PdfPCell main_title = new PdfPCell(new Paragraph("Bakliwal Tutorials",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 32, BaseColor.WHITE)));
            main_title.setBackgroundColor(BaseColor.BLACK);
            main_title.setHorizontalAlignment(Element.ALIGN_CENTER);
            main_title.setVerticalAlignment(Element.ALIGN_MIDDLE);
            main_title.setBorderColor(BaseColor.BLACK);
            main_title.setPaddingBottom(5);
            main_heading_table.addCell(main_title);
            
            PdfPCell emptyCell=new PdfPCell();
            emptyCell.setBackgroundColor(BaseColor.BLACK);
            emptyCell.setBorderColor(BaseColor.BLACK);
            main_heading_table.addCell(emptyCell);
            
            chapter_mark_analysis.add(main_heading_table);
            chapter_mark_analysis.add(new Paragraph("\n"));
            
            PdfPTable secondary_heading_table = new PdfPTable(1);
            secondary_heading_table.setWidthPercentage(100);
            
            PdfPCell student_info = new PdfPCell(new Paragraph(currdata.get("student_roll_no") + " - " + currdata.get("name"),
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.WHITE)));
            student_info.setBackgroundColor(new BaseColor(0, 0, 128));
            student_info.setHorizontalAlignment(Element.ALIGN_CENTER);
            student_info.setVerticalAlignment(Element.ALIGN_MIDDLE);
            student_info.setBorderColor(new BaseColor(0, 0, 128));
            student_info.setPaddingBottom(2);
            secondary_heading_table.addCell(student_info);
            
            PdfPCell test_info = new PdfPCell(new Paragraph("Detailed Analysis for " + BTest_Name + " conducted on " + test_date,
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.WHITE)));
            test_info.setBackgroundColor(new BaseColor(0, 0, 128));
            test_info.setHorizontalAlignment(Element.ALIGN_CENTER);
            test_info.setVerticalAlignment(Element.ALIGN_MIDDLE);
            test_info.setBorderColor(new BaseColor(0, 0, 128));
            test_info.setPaddingBottom(5);
            secondary_heading_table.addCell(test_info);
           
            chapter_mark_analysis.add(secondary_heading_table);
            chapter_mark_analysis.add(new Paragraph("\n"));
            
            PdfPTable local_heading_table_1 = new PdfPTable(1);
            local_heading_table_1.setWidthPercentage(100);
            
            PdfPCell marks_analysis_heading = new PdfPCell(new Paragraph("Marks Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            marks_analysis_heading.setBackgroundColor(new BaseColor(207, 226, 243));
            marks_analysis_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
            marks_analysis_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
            marks_analysis_heading.setBorderColor(new BaseColor(0, 0, 128));
            marks_analysis_heading.setPaddingBottom(3);
            local_heading_table_1.addCell(marks_analysis_heading);
            
            Section section_mark_1 = chapter_mark_analysis.addSection(new Paragraph(""));
            section_mark_1.setNumberDepth(0);
            section_mark_1.add(local_heading_table_1);
            section_mark_1.add(new Paragraph("\n"));

            // Add a table for Marks Analysis
            
            float[] columnWidths = {10f, 9f, 9f, 9f, 9f, 9f, 9f, 9f, 9f, 9f};
            PdfPTable table = new PdfPTable(columnWidths);
            table.setWidthPercentage(100);
            com.itextpdf.text.Font defaultFont = FontFactory.getFont(FontFactory.HELVETICA, 8, BaseColor.BLACK);
            float minHeight = 20f;
            // Add table headers
            PdfPCell cell = new PdfPCell(new Paragraph("Subject", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);

            cell = new PdfPCell(new Paragraph("Marks", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Marks/ Avg", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Marks/ 80 percentile", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Positive Marks", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Negative Marks", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Avg Negative Marks", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setPaddingBottom(2);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Total Questions", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Attempted", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Paragraph("Correct", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
           
            double total_marks = 0;
            double total_marks_per_avg = 0;
            double total_positive_marks = 0;
            double total_negative_marks = 0;
            double total_neg_avg = 0;
            int total_qs = 0;
            int total_attempted = 0;
            int total_correct = 0;
            DecimalFormat df = new DecimalFormat("#.##");

            // Add rows
            List<Map<String, Object>> subjectsList = (List<Map<String, Object>>) BTest_data.get("subjects");
            for(int i=0;i<subjectsList.size();i++) {
            	 Map<String, Object> subj_data=subjectsList.get(i);
            	 String subj_name=(String) subj_data.get("subject_name");
             	String subject_name=subjFullForm.get(subj_name);
             	
             	List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subj_data.get("q_types");

                int total_num_qs = qTypes.stream()
                        .mapToInt(qType -> ((Double) qType.get("num_of_qs")).intValue())
                        .sum();
                cell = new PdfPCell(new Paragraph(subject_name, FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                cell.setBackgroundColor(new BaseColor(207, 226, 243));
                table.addCell(cell);
                
                Double subj_total_marks = (Double) currdata.get(subj_name + "_total_marks");
                total_marks += subj_total_marks;
                Double subj_marks_per_avg = (Double) currdata.get("marks_per_avg_" + subj_name);
                total_marks_per_avg += subj_marks_per_avg;
                Double subj_marks_per_80 = (Double) currdata.get("marks_per_80_" + subj_name);
                
                cell = new PdfPCell(new Paragraph(String.valueOf(subj_total_marks.intValue()), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(df.format(subj_marks_per_avg), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(df.format(subj_marks_per_80), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);

                Double subj_pos_marks = (Double) currdata.get(subj_name + "_positive_marks");
                total_positive_marks += subj_pos_marks;
                Double subj_neg_marks = (Double) currdata.get(subj_name + "_negative_marks");
                total_negative_marks += subj_neg_marks;
                Double subj_avg_neg_marks = (Double) neg_averages.get("avg_neg_" + subj_name);
                total_neg_avg += subj_avg_neg_marks;

                cell = new PdfPCell(new Paragraph(String.valueOf(subj_pos_marks.intValue()), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(String.valueOf(subj_neg_marks.intValue()), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(df.format(subj_avg_neg_marks), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);

                total_qs += total_num_qs;
                Integer subj_total_attempted = (Integer) currdata.get(subj_name + "_num_attempted");
                total_attempted += subj_total_attempted;
                Integer subj_correct = (Integer) currdata.get(subj_name + "_num_correct");
                total_correct += subj_correct;

                cell = new PdfPCell(new Paragraph(String.valueOf(total_num_qs), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(String.valueOf(subj_total_attempted), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
                cell = new PdfPCell(new Paragraph(String.valueOf(subj_correct), defaultFont));
                cell.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell.setMinimumHeight(minHeight);
                table.addCell(cell);
            }
            
            double sum_of_averages = averages.values().stream().mapToDouble(Double::doubleValue).sum();
            cell = new PdfPCell(new Phrase("Total", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            cell.setBackgroundColor(new BaseColor(207, 226, 243));
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf((int) total_marks), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(df.format(total_marks / sum_of_averages), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(df.format(total_marks / total_marks_per_80), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf((int) total_positive_marks), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf((int) total_negative_marks), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(df.format(total_neg_avg), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf(total_qs), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf(total_attempted), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            cell = new PdfPCell(new Phrase(String.valueOf(total_correct), defaultFont));
            cell.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            cell.setMinimumHeight(minHeight);
            table.addCell(cell);
            //document.add(table);
            section_mark_1.add(table);
            section_mark_1.add(new Paragraph("\n"));
           
            Section section_mark_2 = chapter_mark_analysis.addSection(new Paragraph(""));
            section_mark_2.setNumberDepth(0);
            
            PdfPTable local_heading_table_2 = new PdfPTable(1);
            local_heading_table_2.setWidthPercentage(100);
            
            PdfPCell Question_wise_heading = new PdfPCell(new Paragraph("Question-Wise Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            Question_wise_heading.setBackgroundColor(new BaseColor(207, 226, 243));
            Question_wise_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
            Question_wise_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
            Question_wise_heading.setBorderColor(new BaseColor(0, 0, 128));
            Question_wise_heading.setPaddingBottom(3);
            local_heading_table_2.addCell(Question_wise_heading);
            
            section_mark_2.add(local_heading_table_2);
            section_mark_2.add(new Paragraph("\n"));
            
            PdfPTable Imgtable = new PdfPTable(2);
            Imgtable.setWidthPercentage(100);
            
            
            Map<String, DefaultCategoryDataset> datasetMap= new HashMap<>();

            for (Map<String, Object> subject_data : subjectsList) {
                String subjectName = (String) subject_data.get("subject_name");
                List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subject_data.get("q_types");

                for (Map<String, Object> qType : qTypes) {
                    String qTypeName = (String) qType.get("q_type_name");
                    //PHY_MCO_attempted_perc

                    double attemptedPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_attempted_perc"); // Example: 55% attempted
                    double correctPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_correct_perc"); // Example: 40% correct
                    double incorrectPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_incorrect_perc"); // Example: 15% partially correct

                    DefaultCategoryDataset dataset = datasetMap.computeIfAbsent(qTypeName, k -> new DefaultCategoryDataset());

                    dataset.addValue(correctPercentage, "Correct", subjectName);
                    dataset.addValue(incorrectPercentage, "Incorrect", subjectName);
                    double partiallyCorrectPercentage = attemptedPercentage - (correctPercentage + incorrectPercentage);
                    if (partiallyCorrectPercentage < 0) {
                        partiallyCorrectPercentage = 0;
                    }
                    dataset.addValue(partiallyCorrectPercentage, "Partially Correct", subjectName);

                }
            }
            
            Map<String, JFreeChart> chartMap = new HashMap<>();

            for (String qTypeName : datasetMap.keySet()) {
                DefaultCategoryDataset dataset = datasetMap.get(qTypeName);

                JFreeChart chart = ChartFactory.createStackedBarChart(
                        qTypeName + " Report",
                        "Subjects",
                        "Percentage (%)",
                        dataset,
                        PlotOrientation.VERTICAL,
                        true,
                        true,
                        false
                );

                CategoryPlot plot = (CategoryPlot) chart.getPlot();

                // Set axis range with padding and tick unit
                ValueAxis rangeAxis = plot.getRangeAxis();
                rangeAxis.setRange(0.0, 105.0); // Adding extra space above 100
                if (rangeAxis instanceof NumberAxis) {
                    NumberAxis numberAxis = (NumberAxis) rangeAxis;
                    numberAxis.setTickUnit(new NumberTickUnit(20.0));
                }

                // Customize chart appearance
                chart.setBackgroundPaint(Color.WHITE);
                plot.setBackgroundPaint(Color.LIGHT_GRAY);
                plot.setDomainGridlinePaint(Color.BLACK);
                plot.setRangeGridlinePaint(Color.BLACK);

                // Set colors for different series
                StackedBarRenderer renderer = (StackedBarRenderer) plot.getRenderer();
                renderer.setSeriesPaint(0, Color.GREEN); // Correct
                renderer.setSeriesPaint(1, Color.RED);   // Incorrect
                renderer.setSeriesPaint(2, Color.YELLOW); // Partially Correct

                // Set title font size
                //chart.setTitle(new TextTitle(qTypeName + " Report", new java.awt.Font("Arial", java.awt.Font.BOLD, 36)));
                TextTitle title = new TextTitle(qTypeName + " Report", new java.awt.Font("Arial", java.awt.Font.BOLD, 44));
                //title.setHorizontalAlignment(HorizontalAlignment.CENTER);

                // Add padding to the title to move it down
                title.setPadding(0, 70, 0, 0);

                chart.setTitle(title);


                // Set axis label font size
                plot.getDomainAxis().setLabelFont(new java.awt.Font("Arial", java.awt.Font.BOLD, 36));
                plot.getRangeAxis().setLabelFont(new java.awt.Font("Arial", java.awt.Font.BOLD, 36));

                // Set tick label font size
                plot.getDomainAxis().setTickLabelFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 36));
                plot.getRangeAxis().setTickLabelFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 36));
                
                plot.setInsets(new RectangleInsets(10, 10, 20, 10));


                // Set legend font size
                //LegendTitle legend = chart.getLegend();
                LegendTitle legend = chart.getLegend();
                legend.setPosition(RectangleEdge.BOTTOM);
                legend.setPadding(0, 90, 0, 0);
                legend.setItemFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 36));
//                if (legend != null) {
//                	legend.setItemFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 36));
//                }

                // Generate a larger image for higher quality
                int maxWidth = 1100;  // Maximum width in pixels
                int maxHeight = 800; // Maximum height in pixels

                // Create ByteArrayOutputStream for image
                ByteArrayOutputStream chartOut = new ByteArrayOutputStream();

                // Write chart as PNG with a high DPI and let chart handle aspect ratio
                ChartUtils.writeChartAsPNG(chartOut, chart, maxWidth, maxHeight);

                // Create Image instance
                Image chartImage = Image.getInstance(chartOut.toByteArray());

                // Adjust image alignment and scale factor
                chartImage.setAlignment(Element.ALIGN_CENTER);
                float originalWidth = chartImage.getScaledWidth();
                float originalHeight = chartImage.getScaledHeight();

                // Desired dimensions
                float targetWidth = 200; // Adjust target width as needed
                float targetHeight = 200; // Adjust target height as needed

                // Calculate scaling percentages
                float widthPercentage = (targetWidth / originalWidth) * 100;
                float heightPercentage = (targetHeight / originalHeight) * 100;

                // Use the smaller percentage to fit the image within the target dimensions
                float scalingPercentage = Math.min(widthPercentage, heightPercentage);

                // Scale the image
                chartImage.scalePercent(scalingPercentage);

                // Add image to the document
                PdfPCell currImage= new PdfPCell(chartImage);
                currImage.setBorder(PdfPCell.NO_BORDER);
                currImage.setPaddingBottom(15);
                currImage.setPaddingTop(15);
                currImage.setPaddingLeft(30);
                currImage.setPaddingRight(30);
                
                Imgtable.addCell(currImage);
            }
            if(datasetMap.size()%2!=0) {
            	PdfPCell extraCell=new PdfPCell(new Paragraph(""));
            	extraCell.setBorder(PdfPCell.NO_BORDER);
            	Imgtable.addCell(extraCell);
            }
            section_mark_2.add(Imgtable);
            document.add(chapter_mark_analysis);
            
            Chapter chapter_5_must = new Chapter(2);
            chapter_5_must.setNumberDepth(0);
            
            Section section_5must_1 = chapter_5_must.addSection(new Paragraph(""));
            section_5must_1.setNumberDepth(0);
         
            PdfPTable local_heading_table_3 = new PdfPTable(1);
            local_heading_table_3.setWidthPercentage(100);
            
            PdfPCell Must_attempt_heading = new PdfPCell(new Paragraph("5 Must Attempt Questions for Each Subject",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            Must_attempt_heading.setBackgroundColor(new BaseColor(207, 226, 243));
            Must_attempt_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
            Must_attempt_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
            Must_attempt_heading.setBorderColor(new BaseColor(0, 0, 128));
            Must_attempt_heading.setPaddingBottom(3);
            local_heading_table_3.addCell(Must_attempt_heading);
            
            section_5must_1.add(local_heading_table_3);
            section_5must_1.add(new Paragraph("\n"));
            
            minHeight=15f;
            
            PdfPTable must5_table = new PdfPTable(2);
            must5_table.setWidthPercentage(100);
            
            String set_no = (String) currdata.get("set_no");

            for (int i = 0; i < subjectsList.size(); i++) {
                Map<String, Object> subj_data = subjectsList.get(i);
                String subj_name = (String) subj_data.get("subject_name");
                String subject_name = subjFullForm.get(subj_name);
                
                PdfPTable subj_heading_table = new PdfPTable(1);
                //subj_heading_table.setWidthPercentage(100);
                
                PdfPCell subj_heading = new PdfPCell(new Paragraph(subject_name,
                        FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.WHITE)));
                subj_heading.setBackgroundColor(new BaseColor(0, 0, 128));
                subj_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
                subj_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
                subj_heading.setBorder(PdfPCell.NO_BORDER);
                subj_heading.setPaddingBottom(3);
                subj_heading_table.addCell(subj_heading);
                
                PdfPTable table_5ques = new PdfPTable(3);
                
                PdfPCell cell1 = new PdfPCell(new Phrase("Question Number", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell1.setBackgroundColor(new BaseColor(207, 226, 243));
                cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell1.setMinimumHeight(minHeight);
                table_5ques.addCell(cell1);

                PdfPCell cell2 = new PdfPCell(new Phrase("Attempted Percent", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell2.setBackgroundColor(new BaseColor(207, 226, 243));
                cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell2.setMinimumHeight(minHeight);
                table_5ques.addCell(cell2);

                PdfPCell cell3 = new PdfPCell(new Phrase("Correct Percent", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell3.setBackgroundColor(new BaseColor(207, 226, 243));
                cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell3.setMinimumHeight(minHeight);
                table_5ques.addCell(cell3);

                Map<String, Map<String, Double>> ques_perc = (ques_analysis.get(subj_name)).get(set_no);

                for (int j = 0; j < Math.min(5, ques_perc.size()); j++) {
                    String key = new ArrayList<>(ques_perc.keySet()).get(j);
                    int q_no = Integer.parseInt(key.split("_")[1]);

                    BaseColor backgroundColor = BaseColor.WHITE;
                    String ques_status = (String) currdata.get("status_" + subj_name + "_" + q_no);
                    if ("CORRECT".equals(ques_status)) backgroundColor = new BaseColor(144, 238, 144); // Light green
                    else if ("NOT CORRECT".equals(ques_status)) backgroundColor = new BaseColor(255, 182, 193); // Light red (pink)
                    else if ("PARTIALLY CORRECT".equals(ques_status)) backgroundColor = new BaseColor(255, 235, 128); // gold yellow

                    Map<String, Double> stats = ques_perc.get(key);
                    Double attempted_perc = stats.get("attempted_perc");
                    Double correct_perc = stats.get("correct_perc");

                    PdfPCell cell4 = new PdfPCell(new Phrase(String.valueOf(q_no), defaultFont));
                    cell4.setBackgroundColor(backgroundColor);
                    cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell4.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell4.setMinimumHeight(minHeight);
                    table_5ques.addCell(cell4);

                    PdfPCell cell5 = new PdfPCell(new Phrase(df.format(attempted_perc), defaultFont));
                    cell5.setBackgroundColor(backgroundColor);
                    cell5.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell5.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell5.setMinimumHeight(minHeight);
                    table_5ques.addCell(cell5);

                    PdfPCell cell6 = new PdfPCell(new Phrase(df.format(correct_perc), defaultFont));
                    cell6.setBackgroundColor(backgroundColor);
                    cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell6.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell6.setMinimumHeight(minHeight);
                    table_5ques.addCell(cell6);
                }
                PdfPCell currsubjecttable=new PdfPCell();
                currsubjecttable.addElement(subj_heading_table);
                currsubjecttable.addElement(table_5ques);
                currsubjecttable.setPaddingBottom(20);
                currsubjecttable.setPaddingTop(20);
                currsubjecttable.setBorder(PdfPCell.NO_BORDER);
                must5_table.addCell(currsubjecttable);
            }
            if(subjectsList.size()%2!=0) {
            	PdfPCell extraCell=new PdfPCell(new Paragraph("")); 
            	extraCell.setBorder(PdfPCell.NO_BORDER);
            	must5_table.addCell(extraCell);
            }
            section_5must_1.add(must5_table);
            document.add(chapter_5_must);
            
            Chapter chapter_individual_ques = new Chapter(3);
            chapter_individual_ques.setNumberDepth(0);
            //document.add(local_heading_4);
            
            Section section_individual_1 = chapter_individual_ques.addSection(new Paragraph(""));
            section_individual_1.setNumberDepth(0);
            
            PdfPTable local_heading_table_4 = new PdfPTable(1);
            local_heading_table_4.setWidthPercentage(100);
            
            PdfPCell individual_ques_heading = new PdfPCell(new Paragraph("Individual Question Wise Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            individual_ques_heading.setBackgroundColor(new BaseColor(207, 226, 243));
            individual_ques_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
            individual_ques_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);;
            individual_ques_heading.setBorderColor(new BaseColor(0, 0, 128));
            individual_ques_heading.setPaddingBottom(3);
            local_heading_table_4.addCell(individual_ques_heading);
            
            section_individual_1.add(local_heading_table_4);
            section_individual_1.add(new Paragraph("\n"));
            
            PdfPTable table_individual = new PdfPTable(2);
            table_individual.setWidthPercentage(100);
            
            
            for (int i = 0; i < subjectsList.size(); i++) {
                Map<String, Object> subj_data = subjectsList.get(i);
                String subj_name = (String) subj_data.get("subject_name");
                String subject_name = subjFullForm.get(subj_name);
                List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subj_data.get("q_types");

                int total_num_qs = qTypes.stream()
                    .mapToInt(qType -> ((Double) qType.get("num_of_qs")).intValue())
                    .sum();

                PdfPTable table_individual_ques = new PdfPTable(3);
                
                PdfPTable subj_heading_table = new PdfPTable(1);
                PdfPCell subj_heading = new PdfPCell(new Paragraph(subject_name,
                        FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.WHITE)));
                subj_heading.setBackgroundColor(new BaseColor(0, 0, 128));
                subj_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
                subj_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
                subj_heading.setBorder(PdfPCell.NO_BORDER);
                subj_heading.setPaddingBottom(3);
                subj_heading_table.addCell(subj_heading);
            
                PdfPCell cell1 = new PdfPCell(new Phrase("Question Number", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell1.setBackgroundColor(new BaseColor(207, 226, 243));
                cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell1.setMinimumHeight(minHeight);
                table_individual_ques .addCell(cell1);

                PdfPCell cell2 = new PdfPCell(new Phrase("Attempted Percent", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell2.setBackgroundColor(new BaseColor(207, 226, 243));
                cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell2.setMinimumHeight(minHeight);
                table_individual_ques .addCell(cell2);

                PdfPCell cell3 = new PdfPCell(new Phrase("Correct Percent", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.BLACK)));
                cell3.setBackgroundColor(new BaseColor(207, 226, 243));
                cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
                cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);
                cell3.setMinimumHeight(minHeight);
                table_individual_ques .addCell(cell3);

                for (int j = 1; j <= total_num_qs; j++) {

                    Map<String, Map<String, Double>> ques_perc = (ques_analysis.get(subj_name)).get(set_no);
                    Map<String, Double> stats = ques_perc.get(subj_name + "_" + j);

                    BaseColor backgroundColor = BaseColor.WHITE;
                    String ques_status = (String) currdata.get("status_" + subj_name + "_" + j);
                    if ("CORRECT".equals(ques_status)) backgroundColor = new BaseColor(144, 238, 144); // Light green
                    else if ("NOT CORRECT".equals(ques_status)) backgroundColor = new BaseColor(255, 182, 193); // Light red (pink)
                    else if ("PARTIALLY CORRECT".equals(ques_status)) backgroundColor = new BaseColor(255, 235, 128); // Gold yellow

                    Double attempted_perc = stats.get("attempted_perc");
                    Double correct_perc = stats.get("correct_perc");

                    PdfPCell cell4 = new PdfPCell(new Phrase(String.valueOf(j), defaultFont));
                    cell4.setBackgroundColor(backgroundColor);
                    cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell4.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell4.setMinimumHeight(minHeight);
                    table_individual_ques .addCell(cell4);

                    PdfPCell cell5 = new PdfPCell(new Phrase(df.format(attempted_perc), defaultFont));
                    cell5.setBackgroundColor(backgroundColor);
                    cell5.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell5.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell5.setMinimumHeight(minHeight);
                    table_individual_ques .addCell(cell5);

                    PdfPCell cell6 = new PdfPCell(new Phrase(df.format(correct_perc), defaultFont));
                    cell6.setBackgroundColor(backgroundColor);
                    cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
                    cell6.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    cell6.setMinimumHeight(minHeight);
                    table_individual_ques .addCell(cell6);
                }
                PdfPCell currsubjecttable=new PdfPCell();
                currsubjecttable.addElement(subj_heading_table);
                currsubjecttable.addElement(table_individual_ques);
                currsubjecttable.setPaddingBottom(20);
                currsubjecttable.setPaddingTop(20);
                currsubjecttable.setBorder(PdfPCell.NO_BORDER);
                table_individual.addCell(currsubjecttable);
            }
            if(subjectsList.size()%2!=0) {
            	PdfPCell extraCell=new PdfPCell(new Paragraph("")); 
            	extraCell.setBorder(PdfPCell.NO_BORDER);
            	table_individual.addCell(extraCell);
            }
            section_individual_1.add(table_individual);
            document.add(chapter_individual_ques);
            
            
            PdfPTable local_heading_table_5 = new PdfPTable(1);
            local_heading_table_5.setWidthPercentage(100);
            
            PdfPCell interpretion_heading = new PdfPCell(new Paragraph("How to Interpret this Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            interpretion_heading.setBackgroundColor(new BaseColor(207, 226, 243));
            interpretion_heading.setHorizontalAlignment(Element.ALIGN_CENTER);
            interpretion_heading.setVerticalAlignment(Element.ALIGN_MIDDLE);
            interpretion_heading.setBorderColor(new BaseColor(0, 0, 128));
            interpretion_heading.setPaddingBottom(3);
            local_heading_table_5.addCell(interpretion_heading);
            
            Chapter chapter_iterpretion = new Chapter(4);
            chapter_iterpretion.setNumberDepth(0);
            chapter_iterpretion.add(local_heading_table_5);
            chapter_iterpretion.add(new Paragraph("\n"));
            
         // Load the images
            InputStream markAnalysisStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Mark_Analysis_BTest.jpg");
            InputStream quesTypeAnalysisStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Ques_Type_Analysis.jpg");
            InputStream mustAttemptAnalysisStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/5_Must_Attempt.png");
            
            // Convert the InputStream to byte array
            byte[] markAnalysisBytes = IOUtils.toByteArray(markAnalysisStream);
            byte[] quesTypeAnalysisBytes = IOUtils.toByteArray(quesTypeAnalysisStream);
            byte[] mustAttemptAnalysisBytes = IOUtils.toByteArray(mustAttemptAnalysisStream);


            Section section_interpretion_1 = chapter_iterpretion.addSection(new Paragraph("\n"));
            section_interpretion_1.setNumberDepth(0);
            
            PdfPTable marksImage_heading_table = new PdfPTable(1);
            marksImage_heading_table.setWidthPercentage(100);
            
            PdfPCell heading_mark_analysis = new PdfPCell(new Paragraph("Mark Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            heading_mark_analysis.setBackgroundColor(new BaseColor(207, 226, 243));
            heading_mark_analysis.setHorizontalAlignment(Element.ALIGN_CENTER);
            heading_mark_analysis.setVerticalAlignment(Element.ALIGN_MIDDLE);
            heading_mark_analysis.setBorderColor(new BaseColor(0, 0, 128));
            heading_mark_analysis.setPaddingBottom(3);
            marksImage_heading_table.addCell(heading_mark_analysis);
            
            section_interpretion_1.add(marksImage_heading_table);
            section_interpretion_1.add(new Paragraph("\n"));
            
            Image markAnalysisImage = Image.getInstance(markAnalysisBytes);
           // markAnalysisImage.setAbsolutePosition(36, 600); // Positioning the image on the page
            markAnalysisImage.scaleToFit(500, 300);
            section_interpretion_1.add(markAnalysisImage);
            section_interpretion_1.add(new Paragraph("\n"));
            
            Section section_interpretion_2 = chapter_iterpretion.addSection(new Paragraph(""));
            section_interpretion_2.setNumberDepth(0);
            
            PdfPTable questype_heading_table = new PdfPTable(1);
            questype_heading_table.setWidthPercentage(100);
            
            PdfPCell ques_type_analysis = new PdfPCell(new Paragraph("Question Type Analysis",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            ques_type_analysis.setBackgroundColor(new BaseColor(207, 226, 243));
            ques_type_analysis.setHorizontalAlignment(Element.ALIGN_CENTER);
            ques_type_analysis.setVerticalAlignment(Element.ALIGN_MIDDLE);
            ques_type_analysis.setBorderColor(new BaseColor(0, 0, 128));
            ques_type_analysis.setPaddingBottom(3);
            questype_heading_table.addCell(ques_type_analysis);
            
            section_interpretion_2.add(questype_heading_table);
            section_interpretion_2.add(new Paragraph("\n"));

            // Add the Question Type Analysis image
            Image quesTypeAnalysisImage = Image.getInstance(quesTypeAnalysisBytes);
            //quesTypeAnalysisImage.setAbsolutePosition(36, 450); // Positioning the image on the page
            quesTypeAnalysisImage.scaleToFit(500, 300);
            section_interpretion_2.add(quesTypeAnalysisImage);
            section_interpretion_2.add(new Paragraph("\n"));
            //document.add(quesTypeAnalysisImage);

            // Add another heading before the next image
            Section section_interpretion_3 = chapter_iterpretion.addSection(new Paragraph(""));
            section_interpretion_3.setNumberDepth(0);
            
            PdfPTable mustAttemptImage_table = new PdfPTable(1);
            mustAttemptImage_table .setWidthPercentage(100);
            
            PdfPCell must_attempt_analysis = new PdfPCell(new Paragraph("5 Must Attempt Questions for Each Subject",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK)));
            must_attempt_analysis.setBackgroundColor(new BaseColor(207, 226, 243));
            must_attempt_analysis.setHorizontalAlignment(Element.ALIGN_CENTER);
            must_attempt_analysis.setVerticalAlignment(Element.ALIGN_MIDDLE);
            must_attempt_analysis.setBorderColor(new BaseColor(0, 0, 128));
            must_attempt_analysis.setPaddingBottom(3);
            mustAttemptImage_table .addCell(must_attempt_analysis);
            
            section_interpretion_3.add(mustAttemptImage_table );
            section_interpretion_3.add(new Paragraph("\n"));
            //document.add(mustAttemptHeading);

            // Add the 5 Must Attempt Questions image
            Image mustAttemptAnalysisImage = Image.getInstance(mustAttemptAnalysisBytes);
            //mustAttemptAnalysisImage.setAbsolutePosition(36, 300); // Positioning the image on the page
            mustAttemptAnalysisImage.scaleToFit(500, 300);
            section_interpretion_3.add(mustAttemptAnalysisImage);
            section_interpretion_3.add(new Paragraph("\n"));
            //document.add(mustAttemptAnalysisImage);
            
            document.add(chapter_iterpretion);

            // Close the document
            document.close();
            //System.out.println("PDF Created");

        } catch (IOException | com.itextpdf.text.DocumentException e) {
            e.printStackTrace();
        }
	}
	
	private static XSSFColor blackColor;
	private static XSSFCellStyle style;
	private static XSSFCellStyle localHeadingStyle;
	private static XSSFColor greenColor;;
	private static XSSFCellStyle styleWithGreenBorder;
	private static XSSFColor redColor;
	private static XSSFCellStyle styleWithRedBorder;
	private static XSSFColor yellowColor;
	private static XSSFCellStyle styleWithYellowBorder;
	private static XSSFColor mainheadingGreenColor;
	private static XSSFColor whiteColor;
	private static XSSFColor headingGreenColor;
	private static XSSFCellStyle styleMainHeading;
	private static XSSFCellStyle styleHeading;						
	private static XSSFColor headingBlueColor;
	private static XSSFCellStyle blueStyleHeading;
	
	public static void createReports() {
		try {
		String folderName = BTest_Name;
		 String desktopPath = getDesktopPath();
	        if (desktopPath == null) {
	            throw new IllegalStateException("Desktop path could not be found");
	        }

	        // Create the new folder path
	        Path newFolderPath = Paths.get(desktopPath, folderName);

	        try {
	            // Create the folder if it doesn't exist
	            if (!Files.exists(newFolderPath)) {
	                Files.createDirectory(newFolderPath);
	            }
	            System.out.println("Folder created at: " + newFolderPath.toString());
	            folderFilePath = newFolderPath.toString();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        FileInputStream file = new FileInputStream(new File(excelPath));
	     XSSFWorkbook workbook=new XSSFWorkbook(file);
	     setAllStyles(workbook);
	    int total=finalOutput.size();
	    //int total=10;
	    long stTime = System.currentTimeMillis(); 
	    ZipSecureFile.setMinInflateRatio(0);
		for(int i=0;i<total;i++) {
			//System.out.println(finalOutput.size());
			populateTrialData( workbook, i);
			App.updateProgress(i+1 + "/" + total);
			if(i%50 == 0) {
				try {
				Runtime.getRuntime().gc();
				System.gc();
				file.close();
				workbook.close();
//				Thread.sleep(2000);
				file = new FileInputStream(new File(excelPath));
				workbook=new XSSFWorkbook(file);
				setAllStyles(workbook);
				}
				catch(Exception e) {
					e.printStackTrace();
					
				}
			}
			System.out.println("i: "+i+"\tTime: "+(System.currentTimeMillis()-stTime));
			stTime = System.currentTimeMillis();
		}
//		File file=new File(excelPath);
//		file.delete();
		file.close();
		workbook.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void createWorkbook() throws IOException {
		t2.clear();
//		ConvertFormat.student_marks_arr_1.clear();
//		ConvertFormat.student_info_arr.clear();
//		ConvertFormat.q_info_arr.clear();
//		ConvertFormat.correction_arr.clear();
		XSSFWorkbook resultWorkbook=new XSSFWorkbook();
		String fileName=BTest_Name + ".xlsx";
		String desktopPath = getDesktopPath();
		String filePath = desktopPath + File.separator + fileName;
		 try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
	            resultWorkbook.write(fileOut);
	            System.out.println("Workbook saved to: " + filePath);
	        } catch (IOException e) {
	            System.err.println("Failed to save the workbook. Error: " + e.getMessage());
	            e.printStackTrace();
	        } finally {
	            // Close the workbook
	            try {
	                resultWorkbook.close();
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        }
		 resultWorkbook.close();
		 
		 excelPath=filePath;
		 FileInputStream file = new FileInputStream(new File(excelPath));
		 XSSFWorkbook workbook= new XSSFWorkbook(file);
		 if(!islastSet) {
			 setAllStyles(workbook);
			 populateTrialData(workbook,0);
			 }
		 file.close();
		 workbook.close();
		 String libreOfficePath = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
		 processBuilder = new ProcessBuilder(
		            libreOfficePath,
		            "--headless",
		            "--convert-to",
		            "pdf",
		            "--outdir",
		            new File(excelPath).getParent(),
		            excelPath
		    );
	}

	private static String getDesktopPath() {
	    String userHome = System.getProperty("user.home");
	    Path defaultDesktopPath = Paths.get(userHome, "Desktop");

	    // Check default desktop path first
	    if (Files.exists(defaultDesktopPath)) {
	        return defaultDesktopPath.toString();
	    }

	    // Check for OneDrive path
	    String oneDrivePath = System.getenv("OneDrive");
	    if (oneDrivePath != null && !oneDrivePath.isEmpty()) {
	        Path oneDriveDesktopPath = Paths.get(oneDrivePath, "Desktop");
	        if (Files.exists(oneDriveDesktopPath)) {
	            return oneDriveDesktopPath.toString();
	        }
	    }

	    // If neither path is found, return the user home directory
	    return userHome;
	}

	
	private static void populateTrialData(XSSFWorkbook workbook, Integer ind) {
		try {
			//System.out.println(ind);
			Map<String, Object> currdata=finalOutput.get(ind);
			//FileInputStream file = new FileInputStream(new File(excelFilePath));
			//XSSFWorkbook workbook = new XSSFWorkbook(file);
			//XSSFSheet outputsheet = workbook.createSheet((String)currdata.get("student_roll_no")); 
			//XSSFSheet outputsheet=workbook.getSheet("Trial");
			XSSFSheet outputsheet= workbook.getSheet("Btest Report");
			if(outputsheet!=null && !islastSet) {
				workbook.removeSheetAt(workbook.getSheetIndex(outputsheet));
				outputsheet=null;
			}
			if(outputsheet==null)outputsheet=workbook.createSheet("Btest Report");
			int sheetIndex = workbook.getSheetIndex(outputsheet);
			//setBordersToWhite(outputsheet);
			//setSheetAppearance(outputsheet);
			//System.out.println("inputstream is successful " + ind);
			
	        
//	        Row firstRow=outputsheet.getRow(0);
//	        firstRow.setHeight((short) (3*(1440/2.54f)));
	        if(!islastSet) {
	        	CellRangeAddress newMergedRegion = new CellRangeAddress(0, 0, 0, 9);
	        	if (!isRangeAlreadyMerged(outputsheet, newMergedRegion)) {
	                // Add the merged region if it is not already merged
	                outputsheet.addMergedRegion(newMergedRegion);
	            }
	       // outputsheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 9));
	        Row row = outputsheet.createRow(0);
	        row.setHeight((short) (1.75*(1440/2.54f)));
	        Cell cell = row.createCell(0);
	        cell.setCellValue("Bakliwal Tutorials");
	        XSSFCellStyle stylemain = workbook.createCellStyle();
	        XSSFFont font = workbook.createFont();
	        font.setColor(whiteColor);
	        font.setFontName("Playfair Display");
	        font.setBold(true);
	        font.setFontHeightInPoints((short) (36));
	        stylemain.setAlignment(HorizontalAlignment.CENTER);
	        stylemain.setVerticalAlignment(VerticalAlignment.CENTER);
	        stylemain.setFont(font);
	        stylemain.setFillForegroundColor(blackColor);
	        stylemain.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        cell.setCellStyle(stylemain);
	        for(int i=0;i<=9;i++) {
	        	int columnwidth;
	        	switch(i){
	        	case 0:
	        		columnwidth=(int) (2.75*60*20);
	        		break;
	        	case 1:
	        	case 4:
	        	case 7:
	        		columnwidth=(int) (4*60*20);
	        		break;
	        	case 3:
	        	case 6:
	        		columnwidth=(int) (2*60*20);
	        		break;
	        	default:
	        		columnwidth=(int) (2.5*60*20);
	        		break;
	        	}
	        	outputsheet.setColumnWidth(i, columnwidth);
	        }
	        
	        }
	        //System.out.println("not the heading problem");
	        int pageStartRow=0;
            List<Map<String, Object>> subjectsList = (List<Map<String, Object>>) BTest_data.get("subjects");
           
            mergeAndSetCellValue(outputsheet, 2, 2, 0, 9, currdata.get("student_roll_no") + " - " + currdata.get("name"), styleMainHeading);
            if(!islastSet)mergeAndSetCellValue(outputsheet, 3, 3, 0, 9, "Detailed Analysis for " + BTest_Name + " conducted on " + test_date, styleMainHeading);
            
            if(!islastSet)mergeAndSetCellValue(outputsheet, 5, 5, 0, 9, "Marks Analysis", styleHeading);
	        
	        
//			CellRangeAddress mergedRegion2 = new CellRangeAddress(3, 3, 0, 8);
//	        outputsheet.addMergedRegion(mergedRegion2);
            Double total_marks=0.0;
            Double total_marks_per_avg=0.0;
//            Double total_marks_per_80=0.0;
//            for (Double value : percentiles.values()) {
//                total_marks_per_80+= value;
//            }
            Double total_neg_avg=0.0;
            Double total_negative_marks=0.0;
            Double total_positive_marks=0.0;
            int total_qs=0;
            int total_attempted=0;
            int total_correct=0;
            for(Integer i=0;i<subjectsList.size();i++) {
            	Integer start_row=8+5*i;
            	Integer start_col=0;
            	Map<String, Object> subj_data=subjectsList.get(i);
            	//System.out.println(subj_data);
            	String subj_name=(String) subj_data.get("subject_name");
            	String subject_name=subjFullForm.get(subj_name);
            	List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subj_data.get("q_types");

                int total_num_qs = qTypes.stream()
                        .mapToInt(qType -> ((Double) qType.get("num_of_qs")).intValue())
                        .sum();
                if(!islastSet)setCellValue(outputsheet,start_row,start_col,subject_name,localHeadingStyle);
            	start_row++;
            	start_col++;
            	
            	//System.out.println(subj_name);
            	
            	if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Marks",style);
            	setCellValue(outputsheet,start_row+1,start_col,"Marks/ Avg",style);
            	setCellValue(outputsheet,start_row+2,start_col,"Marks/ 80 percentile",style);}
            	start_col++;
            	
            	Double subj_total_marks=(Double)currdata.get(subj_name + "_total_marks");
            	total_marks+=subj_total_marks;
            	Double subj_marks_per_avg=(Double)currdata.get("marks_per_avg_" + subj_name);
            	total_marks_per_avg+=subj_marks_per_avg;
            	Double subj_marks_per_80=(Double)currdata.get("marks_per_80_" + subj_name);
            	//total_marks_per_80+=subj_marks_per_80;
            	//System.out.println(subj_name);
            	setCellValue(outputsheet,start_row,start_col,subj_total_marks,style);
            	setCellValue(outputsheet,start_row+1,start_col,df.format(subj_marks_per_avg),style);
            	setCellValue(outputsheet,start_row+2,start_col,df.format(subj_marks_per_80),style);
            	start_col+=2;
            	//System.out.println(subj_name);
            	
            	
            	if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Positive Marks",style);
            	setCellValue(outputsheet,start_row+1,start_col,"Negative Marks",style);
            	setCellValue(outputsheet,start_row+2,start_col,"Avg Negative Marks",style);}
            	start_col++;
            	
            	Double subj_pos_marks=(Double)currdata.get(subj_name + "_positive_marks");
            	total_positive_marks+=subj_pos_marks;
            	Double subj_neg_marks=(Double)currdata.get(subj_name + "_negative_marks");
            	total_negative_marks+=subj_neg_marks;
            	Double subj_avg_neg_marks=(Double)neg_averages.get("avg_neg_" + subj_name);
            	total_neg_avg+=subj_avg_neg_marks;
            	
            	setCellValue(outputsheet,start_row,start_col,subj_pos_marks,style);
            	setCellValue(outputsheet,start_row+1,start_col,subj_neg_marks,style);
            	setCellValue(outputsheet,start_row+2,start_col,df.format(subj_avg_neg_marks),style);
            	start_col+=2;
            	//System.out.println(subj_name);
            	
            	
            	if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Total Questions",style);
            	setCellValue(outputsheet,start_row+1,start_col,"Attempted",style);
            	setCellValue(outputsheet,start_row+2,start_col,"Correct",style);}
            	start_col++;
            	
            	total_qs+=total_num_qs;
            	Integer subj_total_attempted=(Integer)currdata.get(subj_name + "_num_attempted");
            	total_attempted+=subj_total_attempted;
            	Integer subj_correct=(Integer)currdata.get(subj_name + "_num_correct");
            	total_correct+=subj_correct;
            	
             	setCellValue(outputsheet,start_row,start_col,total_num_qs,style);
            	setCellValue(outputsheet,start_row+1,start_col,subj_total_attempted,style);
            	setCellValue(outputsheet,start_row+2,start_col,subj_correct,style);
            	//System.out.println(subj_name);
            }
            //System.out.println("for loop completes");
            // for total now
            
            double sum_of_averages = averages.values().stream().mapToDouble(Double::doubleValue).sum();
            int start_row=8+5*subjectsList.size();
            int start_col=0;
            if(!islastSet)setCellValue(outputsheet,start_row,start_col,"Total",localHeadingStyle);
            start_row++;
            start_col++;
            //System.out.println("1");
            if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Marks",style);
        	setCellValue(outputsheet,start_row+1,start_col,"Marks/ Avg",style);
        	setCellValue(outputsheet,start_row+2,start_col,"Marks/ 80 percentile",style);}
        	start_col++;
        	setCellValue(outputsheet,start_row,start_col,total_marks,style);
        	setCellValue(outputsheet,start_row+1,start_col,df.format(total_marks/sum_of_averages),style);
        	setCellValue(outputsheet,start_row+2,start_col,df.format(total_marks/total_marks_per_80),style);
        	start_col+=2;
        	//System.out.println("2");
        	if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Positive Marks",style);
        	setCellValue(outputsheet,start_row+1,start_col,"Negative Marks",style);
        	setCellValue(outputsheet,start_row+2,start_col,"Avg Negative Marks",style);}
        	start_col++;
        	setCellValue(outputsheet,start_row,start_col,total_positive_marks,style);
        	setCellValue(outputsheet,start_row+1,start_col,total_negative_marks,style);
        	setCellValue(outputsheet,start_row+2,start_col,df.format(total_neg_avg),style);
        	start_col+=2;
        	//System.out.println("3");
        	if(!islastSet) {setCellValue(outputsheet,start_row,start_col,"Total Questions",style);
        	setCellValue(outputsheet,start_row+1,start_col,"Attempted",style);
        	setCellValue(outputsheet,start_row+2,start_col,"Correct",style);}
        	start_col++;
        	setCellValue(outputsheet,start_row,start_col,total_qs,style);
        	setCellValue(outputsheet,start_row+1,start_col,total_attempted,style);
        	setCellValue(outputsheet,start_row+2,start_col,total_correct,style);
        	
        	start_row+=5;//crossing the total
        	
        	//start_row+=1;//making some gap
//        	if(!islastSet) {
//           	 //workbook.setPrintArea(sheetIndex, 0, 9, pageStartRow, start_row-1);
//           	 outputsheet.setRowBreak(start_row-1);
//                pageStartRow=start_row; 
//            }
        	if(!islastSet)mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "Question Type Analysis", styleHeading);
        	
        	//here heading will go
        	
        	start_row+=3;
        	
        	 Map<String, DefaultCategoryDataset> datasetMap= new HashMap<>();

             for (Map<String, Object> subject_data : subjectsList) {
                 String subjectName = (String) subject_data.get("subject_name");
                 List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subject_data.get("q_types");

                 for (Map<String, Object> qType : qTypes) {
                     String qTypeName = (String) qType.get("q_type_name");
                     //PHY_MCO_attempted_perc

                     double attemptedPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_attempted_perc"); // Example: 55% attempted
                     double correctPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_correct_perc"); // Example: 40% correct
                     double incorrectPercentage = (double) currdata.get(subjectName + "_" + qTypeName + "_incorrect_perc"); // Example: 15% partially correct

                     DefaultCategoryDataset dataset = datasetMap.computeIfAbsent(qTypeName, k -> new DefaultCategoryDataset());

                     dataset.addValue(correctPercentage, "Correct", subjectName);
                     dataset.addValue(incorrectPercentage, "Incorrect", subjectName);
                     double partiallyCorrectPercentage = attemptedPercentage - (correctPercentage + incorrectPercentage);
                     if (partiallyCorrectPercentage < 0) {
                         partiallyCorrectPercentage = 0;
                     }
                     dataset.addValue(partiallyCorrectPercentage, "Partially Correct", subjectName);

                 }
             }
             
             Drawing prevdrawing = outputsheet.getDrawingPatriarch();
             if (prevdrawing instanceof XSSFDrawing) {
            	   for (XSSFShape shape : ((XSSFDrawing)prevdrawing).getShapes()) {
            	    if (shape instanceof XSSFPicture) {
            	     XSSFPicture xssfPicture = (XSSFPicture)shape;
            	     String shapename = xssfPicture.getShapeName();
            	     int row = xssfPicture.getClientAnchor().getRow1();
            	     int col = xssfPicture.getClientAnchor().getCol1();
            	     //System.out.println("Picture " + "" + " with Shapename: " + shapename + " is located row: " + row + ", col: " + col);

            	     if (shapename.contains("Picture")) {
            	    	 deleteEmbeddedXSSFPicture(xssfPicture);
            	    	 deleteCTAnchor(xssfPicture);
            	     }

            	    }
            	   }
            	  }
           

             Map<String, JFreeChart> chartMap = new HashMap<>();

             for (String qTypeName : datasetMap.keySet()) {
                 DefaultCategoryDataset dataset = datasetMap.get(qTypeName);

                 JFreeChart chart = ChartFactory.createStackedBarChart(
                         qTypeName + " Report",
                         "Subjects",
                         "Percentage (%)",
                         dataset,
                         PlotOrientation.VERTICAL,
                         true,
                         true,
                         false
                 );

                 CategoryPlot plot = (CategoryPlot) chart.getPlot();
                 StackedBarRenderer renderer = new StackedBarRenderer();
                 renderer.setSeriesPaint(0, Color.GREEN);
                 renderer.setSeriesPaint(1, Color.RED);
                 renderer.setSeriesPaint(2, Color.YELLOW);
                 plot.setRenderer(renderer);

                 chartMap.put(qTypeName, chart);
             }

             int rowIndex = start_row;
             int colIndex = 1;

             for (String qTypeName : chartMap.keySet()) {
            	 if(colIndex>=10) {
                	 start_row+=11;
                	 rowIndex+=11;
                	 colIndex=1;
                 }
                 JFreeChart chart = chartMap.get(qTypeName);

                 // Convert chart to image
                 byte[] imageBytes = ChartUtils.encodeAsPNG(chart.createBufferedImage(400, 300));

                 // Add image to the workbook
                 int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                 XSSFDrawing drawing = (XSSFDrawing) outputsheet.createDrawingPatriarch();
                 XSSFClientAnchor anchor = new XSSFClientAnchor();
                 anchor.setCol1(colIndex);
                 anchor.setRow1(rowIndex);

                 // Create the picture
                 XSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
                 picture.resize(0.7); // Adjust the scale factor as needed

                 // Increment colIndex for the next chart
                 colIndex += 3;
             }
             
             //updateOrCreateCharts(workbook, outputsheet, datasetMap);
             start_row+=12;
             if(!islastSet) {
            	 //workbook.setPrintArea(sheetIndex, 0, 9, pageStartRow, start_row-1);
            	 outputsheet.setRowBreak(start_row-1);
                 pageStartRow=start_row; 
             }
             if(!islastSet) mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "5 Must Attempt Questions for Each Subject", styleHeading);
             
             
             start_row+=2;
             start_col=1;
             Integer running_row=start_row;
             
             
             
             String set_no=(String) currdata.get("set_no");
             for(int i=0;i<subjectsList.size();i++) {
            	 Map<String, Object> subj_data=subjectsList.get(i);
            	 String subj_name=(String) subj_data.get("subject_name");
             	String subject_name=subjFullForm.get(subj_name);
             	
             	if(start_col>=10) {
             		start_col=1;
             		start_row=running_row;
             		start_row+=2;
             	}
             	
             	Integer curr_row=start_row;
             	
             	if(!islastSet)setCellValue(outputsheet,curr_row,start_col,subject_name,localHeadingStyle);
            	curr_row+=2;
            	//System.out.println("Working 1");
            	
            	Map<String, Map<String, Double>> ques_perc=(ques_analysis.get(subj_name)).get(set_no);
//            	System.out.println("Working here");
//            	System.out.println(ques_perc.size());
            	for(int j=0;j<Math.min(5,ques_perc.size());j++) {
            		String key = new ArrayList<>(ques_perc.keySet()).get(j);
                	//System.out.println("Working 2");
            	    // Extract the question number from the key
            	    int q_no = Integer.parseInt(key.split("_")[1]);
            	    
            	    XSSFCellStyle currstyle=style;
            	    String ques_status = (String) currdata.get("status_" + subj_name + "_" + q_no);
            	    //System.out.println(ques_status);
            	    if ("CORRECT".equals(ques_status)) currstyle = styleWithGreenBorder;
                    else if ("NOT CORRECT".equals(ques_status)) currstyle = styleWithRedBorder;
                 	else if ("PARTIALLY CORRECT".equals(ques_status)) currstyle = styleWithYellowBorder;
            	    
            	    Map<String, Double> stats = ques_perc.get(key);
            	    Double attempted_perc =stats.get("attempted_perc");
            	    Double correct_perc = stats.get("correct_perc");
            	    //System.out.println("Working 4");
            	    Integer curr_col=start_col;
            	    setCellValue(outputsheet,curr_row,curr_col,"Question Number",currstyle);
                	setCellValue(outputsheet,curr_row+1,curr_col,"Attempted Percent",currstyle);
                	setCellValue(outputsheet,curr_row+2,curr_col,"Correct Percent",currstyle);
                	//curr_col++;
                	setCellValue(outputsheet,curr_row,curr_col+1,q_no,currstyle);
                	setCellValue(outputsheet,curr_row+1,curr_col+1,df.format(attempted_perc),currstyle);
                	setCellValue(outputsheet,curr_row+2,curr_col+1,df.format(correct_perc),currstyle);
                	curr_row+=4;
            	}
            	running_row=Math.max(running_row, curr_row);
            	start_col+=3;
             }
             start_row+=23;//5 ques * 4rows each + 1 heading
             
             if(!islastSet) {
            	 //workbook.setPrintArea(sheetIndex, 0, 9, pageStartRow, start_row-1);
            	 outputsheet.setRowBreak(start_row-1);
                 pageStartRow=start_row; 
             }
             if(!islastSet)mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "Individual Question-wise Analysis", styleHeading);
             
             start_row+=2;
             start_col=1;
             running_row=start_row;
             
             
             //report for all question of the students
             for(int i=0;i<subjectsList.size();i++) {
            	 //System.out.println(i);
            	 Map<String, Object> subj_data=subjectsList.get(i);
             	//System.out.println(subj_data);
             	String subj_name=(String) subj_data.get("subject_name");
             	String subject_name=subjFullForm.get(subj_name);
             	List<Map<String, Object>> qTypes = (List<Map<String, Object>>) subj_data.get("q_types");

                 int total_num_qs = qTypes.stream()
                         .mapToInt(qType -> ((Double) qType.get("num_of_qs")).intValue())
                         .sum();
                 //System.out.println(total_num_qs);
                 if(start_col>=10) {
              		start_col=1;
              		start_row=running_row;
              		start_row+=2;
              	}
                 int curr_row=start_row;
                 int curr_col=start_col;
                 
                 if(curr_row+4-pageStartRow + 1 > 70 && !islastSet) /*75 are the maximum no. of rows in a page*/{
            		 pageStartRow=curr_row;
            		 outputsheet.setRowBreak(curr_row-1);
            	 }
                 
                 if(!islastSet)setCellValue(outputsheet, curr_row, curr_col, subject_name, localHeadingStyle);
                 curr_row+=2;
                 for(int j=1;j<=total_num_qs;j++) {
                	 
                	 if(curr_row+2-pageStartRow + 1 > 70 && !islastSet) /*75 are the maximum no. of rows in a page*/{
                		 pageStartRow=curr_row;
                		 outputsheet.setRowBreak(curr_row-1);
                	 }
                	 
                 	Map<String, Map<String, Double>> ques_perc=(ques_analysis.get(subj_name)).get(set_no);
                 	Map<String, Double> stats = ques_perc.get(subj_name + "_" + j);
                 	
                 	XSSFCellStyle currstyle=style;
                 	String ques_status = (String) currdata.get("status_" + subj_name + "_" + j);
                 	if ("CORRECT".equals(ques_status)) currstyle = styleWithGreenBorder;
                    else if ("NOT CORRECT".equals(ques_status)) currstyle = styleWithRedBorder;
                 	else if ("PARTIALLY CORRECT".equals(ques_status)) currstyle = styleWithYellowBorder;
                 	
                 	Double attempted_perc =stats.get("attempted_perc");
            	    Double correct_perc = stats.get("correct_perc");
            	    setCellValue(outputsheet,curr_row,curr_col,"Question Number",currstyle);
                	setCellValue(outputsheet,curr_row+1,curr_col,"Attempted Percent",currstyle);
                	setCellValue(outputsheet,curr_row+2,curr_col,"Correct Percent",currstyle);
                	//curr_col++;
                	setCellValue(outputsheet,curr_row,curr_col+1,j,currstyle);
                	setCellValue(outputsheet,curr_row+1,curr_col+1,df.format(attempted_perc),currstyle);
                	setCellValue(outputsheet,curr_row+2,curr_col+1,df.format(correct_perc),currstyle);
                	curr_row+=4;
                	running_row=Math.max(running_row, curr_row);
                 }
                 start_col+=3;
                 //System.out.println("coming out of for loop");
             }
             start_row=running_row;
             start_row+=20;
             start_col=0;
             if(!islastSet) {
            	 //workbook.setPrintArea(sheetIndex, 0, 9, pageStartRow, start_row-1);
            	 outputsheet.setRowBreak(start_row-1);
                 pageStartRow=start_row; 
             }
             
             
             if(!islastSet) {
            	 //islastSet=true;
            	 //System.out.println("trying to set");
            	 mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "How to Interpret this analysis", blueStyleHeading);
            	 start_row+=2;
            	 mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "Marks Analysis", blueStyleHeading);
            	 start_row+=2;
            	 
     	        InputStream logo = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Bakliwal_Logo.jpg");
           	    byte[] byteslogo = IOUtils.toByteArray(logo);
                int picturelogo = workbook.addPicture(byteslogo, Workbook.PICTURE_TYPE_JPEG);
                Drawing<?> drawing = outputsheet.createDrawingPatriarch();
                CreationHelper helper_logo = workbook.getCreationHelper();
                ClientAnchor anchor_logo = helper_logo.createClientAnchor();
                anchor_logo.setCol1(0);
                anchor_logo.setRow1(0);
//                anchor_mark_analysis.setCol2();
//                anchor_mark_analysis.setRow2();
                Picture pictlogo = drawing.createPicture(anchor_logo, picturelogo);
                pictlogo.resize();
                ((XSSFPicture) pictlogo).getCTPicture().getNvPicPr().getCNvPr().setName("Bakliwal_Logo");
            	 
            	 InputStream mark_analysisStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Mark_Analysis_BTest.jpg");
            	 byte[] bytesMark_Analysis = IOUtils.toByteArray(mark_analysisStream);
                 int pictureMark_analysis = workbook.addPicture(bytesMark_Analysis, Workbook.PICTURE_TYPE_JPEG);
                 //Drawing<?> drawing = outputsheet.createDrawingPatriarch();
                 CreationHelper helper_mark_analysis = workbook.getCreationHelper();
                 ClientAnchor anchor_mark_analysis = helper_mark_analysis.createClientAnchor();
                 anchor_mark_analysis.setCol1(start_col);
                 anchor_mark_analysis.setRow1(start_row);
                 anchor_mark_analysis.setCol2(start_col+8);
                 anchor_mark_analysis.setRow2(start_row+18);
                 Picture pictMark_analysis = drawing.createPicture(anchor_mark_analysis, pictureMark_analysis);
                 pictMark_analysis.resize(1.2,1.2);
                 ((XSSFPicture) pictMark_analysis).getCTPicture().getNvPicPr().getCNvPr().setName("Mark_Analysis_BTest");
                 
                 start_row+=28;
                 mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "Question Type Analysis", blueStyleHeading);
                 start_row+=2;
                 
                 InputStream ques_type_analysis = Generating_Report.class.getClassLoader().getResourceAsStream("Images/Ques_Type_Analysis.jpg");
            	 byte[] bytesques_type_Analysis = IOUtils.toByteArray(ques_type_analysis);
                 int pictureques_type_analysis = workbook.addPicture(bytesques_type_Analysis, Workbook.PICTURE_TYPE_JPEG);
                 //Drawing<?> drawing = outputsheet.createDrawingPatriarch();
                 CreationHelper helper_ques_type_analysis = workbook.getCreationHelper();
                 ClientAnchor anchor_ques_type_analysis = helper_ques_type_analysis.createClientAnchor();
                 //anchor_ques_type_analysis.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
                 anchor_ques_type_analysis.setCol1(start_col);
                 anchor_ques_type_analysis.setRow1(start_row);
                 anchor_ques_type_analysis.setCol2(start_col+8);
                 anchor_ques_type_analysis.setRow2(start_row+14);
                 Picture pict_ques_type_analysis= drawing.createPicture(anchor_ques_type_analysis, pictureques_type_analysis);
                 pict_ques_type_analysis.resize(1.2,1.0);
                 ((XSSFPicture) pict_ques_type_analysis).getCTPicture().getNvPicPr().getCNvPr().setName("Ques_Type_Analysis");
                 
                 start_row+=20;
                 
                 outputsheet.setRowBreak(start_row-1);
                 mergeAndSetCellValue(outputsheet, start_row, start_row, 0, 9, "5 Must Attempt Questions for Each Subject", blueStyleHeading);
                 start_row+=3;
                 InputStream must_attempt_analysisStream = Generating_Report.class.getClassLoader().getResourceAsStream("Images/5_Must_Attempt.jpg");
            	 byte[] bytesmust_attempt_Analysis = IOUtils.toByteArray(must_attempt_analysisStream);
                 int picturemust_attempt_analysis = workbook.addPicture(bytesmust_attempt_Analysis, Workbook.PICTURE_TYPE_JPEG);
                 //Drawing<?> drawing = outputsheet.createDrawingPatriarch();
                 CreationHelper helper_must_attempt_analysis = workbook.getCreationHelper();
                 ClientAnchor anchor_must_attempt_analysis = helper_must_attempt_analysis.createClientAnchor();
                 anchor_must_attempt_analysis.setCol1(start_col);
                 anchor_must_attempt_analysis.setRow1(start_row);
                 anchor_must_attempt_analysis.setCol2(start_col+8);
                 anchor_must_attempt_analysis.setRow2(start_row+18);
                 Picture pictmust_attempt_analysis = drawing.createPicture(anchor_must_attempt_analysis, picturemust_attempt_analysis);
                 pictmust_attempt_analysis.resize(1.2,0.7);
                 ((XSSFPicture) pictmust_attempt_analysis).getCTPicture().getNvPicPr().getCNvPr().setName("5_Must_Attempt");
                 
                 
                 workbook.setPrintArea(sheetIndex, 0, 9, 0, outputsheet.getLastRowNum()+30);
                 PrintSetup printSetup = outputsheet.getPrintSetup();
                 outputsheet.setFitToPage(true);
                 printSetup.setLandscape(false); 
                 printSetup.setFitWidth((short) 1); 
                 printSetup.setFitHeight((short) 0);
             }
              
             //printSetup.setScale((short) 80);
             //if(!islastSet)workbook.setSheetHidden(sheetIndex, false);
             workbook.setActiveSheet(sheetIndex);
             //workbook.setSheetHidden(sheetIndex, false);
             if(islastSet) {
             for (Sheet sheet : workbook) {
                 if (!(sheet.getSheetName().equals("Btest Report"))) {
                	 //System.out.println(sheet.getSheetName());
                     workbook.setSheetHidden(workbook.getSheetIndex(sheet), true);
                 }
             }
             }
             
            String pdfFilePath=folderFilePath + "\\" + currdata.get("student_roll_no") + "_" + BTest_Name + ".pdf";
			FileOutputStream fileOut = new FileOutputStream(excelPath);
			workbook.write(fileOut);
			fileOut.close();
			//file.close();
			//workbook.close();
			if(islastSet)exportToPdf(pdfFilePath);
			islastSet=true;
//			if(ind%100==0 && ind!=0) {
//				workbook.removeSheetAt(sheetIndex);
//				workbook.write(fileOut);
//				islastSet=false;
//			}
//			fileOut.close();
		}
		catch(Exception e)
		{
		
		}
	}
	
	
	
	private static void setCellValue(Sheet sheet, int rowIndex, int columnIndex, Object value, XSSFCellStyle style) {
	    Row row = sheet.getRow(rowIndex);
	    if (row == null) {
	        row = sheet.createRow(rowIndex);
	    }
	    Cell cell = row.getCell(columnIndex);
	    if (cell == null) {
	        cell = row.createCell(columnIndex);
	    }

	    if (value instanceof String) {
	        cell.setCellValue((String) value);
	    } else if (value instanceof Double) {
	        cell.setCellValue((Double) value);
	    } else if (value instanceof Integer) {
	        cell.setCellValue((Integer) value);
	    } else if (value instanceof Boolean) {
	        cell.setCellValue((Boolean) value);
	    } else if (value instanceof Date) {
	        cell.setCellValue((Date) value);
	    } else if (value instanceof Calendar) {
	        cell.setCellValue((Calendar) value);
	    } else {
	        cell.setCellValue(value.toString());
	    }
	    
	    if (style != null) {
	        cell.setCellStyle(style);
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
    private static void printMap(Map<String, Double> map) {
        for (Map.Entry<String, Double> entry : map.entrySet()) {
            System.out.println(entry.getKey() + "=" + entry.getValue());
        }
    }
    
    public static Map<String, Double> calculateAverages(Map<String, Map<String, Object>> t2, String subject) {
        return t2.values().stream()
                .filter(aggregatedData -> aggregatedData.containsKey(subject + "_total_marks"))
                .collect(Collectors.groupingBy(
                        e -> "avg_" + subject,
                        Collectors.averagingDouble(e -> (Double) e.get(subject + "_total_marks"))
                ));
    }
    public static Map<String, Double> calculateNegativeAverages(Map<String, Map<String, Object>> t2, String subject) {
        return t2.values().stream()
                .filter(aggregatedData -> aggregatedData.containsKey(subject + "_negative_marks"))
                .collect(Collectors.groupingBy(
                        e -> "avg_neg_" + subject,
                        Collectors.averagingDouble(e -> (Double) e.get(subject + "_negative_marks"))
                ));
    }
    public static Map<String, Double> calculatePercentiles(Map<String, Map<String, Object>> t2, String subject, double percentile) {
        List<Double> marks = t2.values().stream()
                .filter(aggregatedData -> aggregatedData.containsKey(subject + "_total_marks"))
                .map(aggregatedData -> (Double) aggregatedData.get(subject + "_total_marks"))
                .sorted()
                .collect(Collectors.toList());

        int index = (int) Math.ceil(percentile * marks.size()) - 1;
        double percentileValue = marks.get(index);
        total_marks_per_80+=percentileValue;

        Map<String, Double> result = new HashMap<>();
        result.put(subject + "_80th_percentile", percentileValue);
        return result;
    }
    private static Map<String, Double> initializeStats() {
        Map<String, Double> stats = new HashMap<>();
        stats.put("total", 0.0);
        stats.put("attempted", 0.0);
        stats.put("correct", 0.0);
        return stats;
    }
    public static XSSFCellStyle createCellStyle(XSSFWorkbook workbook, boolean bold, HorizontalAlignment alignment, BorderStyle borderStyle, XSSFColor borderColor, XSSFColor textColor, XSSFColor backgroundColor) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(alignment);

        XSSFFont font = workbook.createFont();
        font.setBold(bold);
        if (textColor != null) {
            font.setColor(textColor);
        } else {
            font.setColor(new XSSFColor(java.awt.Color.BLACK, null)); // Default to black if textColor is null
        }
        style.setFont(font);

        // Reset all borders to NONE first
        style.setBorderTop(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);

        if (borderStyle != null) {
            style.setBorderTop(borderStyle);
            style.setBorderBottom(borderStyle);
            style.setBorderLeft(borderStyle);
            style.setBorderRight(borderStyle);
        }

        if (borderColor != null) {
            style.setTopBorderColor(borderColor);
            style.setBottomBorderColor(borderColor);
            style.setLeftBorderColor(borderColor);
            style.setRightBorderColor(borderColor);
        } else {
            // Reset border colors to default if borderColor is null
            style.setTopBorderColor(new XSSFColor(java.awt.Color.WHITE, null));
            style.setBottomBorderColor(new XSSFColor(java.awt.Color.WHITE, null));
            style.setLeftBorderColor(new XSSFColor(java.awt.Color.WHITE, null));
            style.setRightBorderColor(new XSSFColor(java.awt.Color.WHITE, null));
        }

        if (backgroundColor != null) {
            style.setFillForegroundColor(backgroundColor);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        } else {
            style.setFillForegroundColor(new XSSFColor(java.awt.Color.WHITE, null)); // Default to white if backgroundColor is null
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        return style;
    }
    
    private static void mergeAndSetCellValue(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, String value, CellStyle style) {
    	removeExistingMergedRegions(sheet, firstRow, lastRow, firstCol, lastCol);
        // Clear the contents of the cells in the range
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            for (int colIndex = firstCol; colIndex <= lastCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }
                cell.setCellValue(""); // Clear the cell content
                cell.setCellStyle(style); // Apply the style to all cells in the range
            }
        }

        // Merge the cells
        CellRangeAddress mergedRegion = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(mergedRegion);

        // Set the value in the first cell of the merged region
        Row row = sheet.getRow(firstRow);
        Cell cell = row.getCell(firstCol);
        cell.setCellValue(value);
        cell.setCellStyle(style); // Apply the style to the merged cell
    }
    private static void removeExistingMergedRegions(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() <= lastRow && mergedRegion.getLastRow() >= firstRow &&
                mergedRegion.getFirstColumn() <= firstCol && mergedRegion.getLastColumn() >= lastCol) {
                sheet.removeMergedRegion(i);
                i--; // Adjust the index to account for the removed region
            }
        }
    }
    private static String getImageName(org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTPicture pic) {
    	//System.out.println("i am  here");
        XmlCursor cursor = pic.newCursor();
        cursor.selectPath("./*");
        while (cursor.toNextSelection()) {
            if ("nvPicPr".equals(cursor.getName().getLocalPart())) {
                cursor.selectPath("./*");
                while (cursor.toNextSelection()) {
                    if ("cNvPr".equals(cursor.getName().getLocalPart())) {
                        return cursor.getAttributeText(new QName("name"));
                    }
                }
            }
        }
        return null;
    }
    
    public static void deleteCTAnchor(XSSFPicture xssfPicture) {
    	  XSSFDrawing drawing = xssfPicture.getDrawing();
    	  XmlCursor cursor = xssfPicture.getCTPicture().newCursor();
    	  cursor.toParent();
    	  if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor) {
    	   for (int i = 0; i < drawing.getCTDrawing().getTwoCellAnchorList().size(); i++) {
    	    if (cursor.getObject().equals(drawing.getCTDrawing().getTwoCellAnchorArray(i))) {
    	     drawing.getCTDrawing().removeTwoCellAnchor(i);
    	     //System.out.println("TwoCellAnchor for picture " + xssfPicture + " was deleted.");
    	    }
    	   }
    	  } else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor) {
    	   for (int i = 0; i < drawing.getCTDrawing().getOneCellAnchorList().size(); i++) {
    	    if (cursor.getObject().equals(drawing.getCTDrawing().getOneCellAnchorArray(i))) {
    	     drawing.getCTDrawing().removeOneCellAnchor(i);
    	     //System.out.println("OneCellAnchor for picture " + xssfPicture + " was deleted.");
    	    }
    	   }
    	  } else if (cursor.getObject() instanceof org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTAbsoluteAnchor) {
    	   for (int i = 0; i < drawing.getCTDrawing().getAbsoluteAnchorList().size(); i++) {
    	    if (cursor.getObject().equals(drawing.getCTDrawing().getAbsoluteAnchorArray(i))) {
    	     drawing.getCTDrawing().removeAbsoluteAnchor(i);
    	     //System.out.println("AbsoluteAnchor for picture " + xssfPicture + " was deleted.");
    	    }
    	   }
    	  }
    	 }
    
    public static void deleteEmbeddedXSSFPicture(XSSFPicture xssfPicture) {
    	  if (xssfPicture.getCTPicture().getBlipFill() != null) {
    	   if (xssfPicture.getCTPicture().getBlipFill().getBlip() != null) {
    	    if (xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed() != null) {
    	     String rId = xssfPicture.getCTPicture().getBlipFill().getBlip().getEmbed();
    	     XSSFDrawing drawing = xssfPicture.getDrawing();
    	     drawing.getPackagePart().removeRelationship(rId);
    	     drawing.getPackagePart().getPackage().deletePartRecursive(drawing.getRelationById(rId).getPackagePart().getPartName());
    	     //System.out.println("Picture " + xssfPicture + " was deleted.");
    	    }
    	   }
    	  }
    	 }
    public static void exportToPdf(String pdfFilePath ) {
        try {
            // Path to LibreOffice Calc executable
//            String libreOfficePath = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";

            // Command to convert Excel to PDF
//            ProcessBuilder processBuilder = new ProcessBuilder(
//                    libreOfficePath,
//                    "--headless",
//                    "--convert-to",
//                    "pdf",
//                    "--outdir",
//                    new File(excelFilePath).getParent(),
//                    excelFilePath
//            );

            //System.out.println("Executing command: " + processBuilder.command());

            // Start the process
            Process process = processBuilder.start();
            int exitCode = process.waitFor(); // Wait for the process to complete

            if (exitCode == 0) {
                //System.out.println("LibreOffice conversion successful");

                // LibreOffice generates a PDF with the same name as the Excel file in the same directory
                String generatedPdfPath = excelPath.replace(".xlsx", ".pdf");
                Path source = Paths.get(generatedPdfPath);
                Path target = Paths.get(pdfFilePath);

                // Rename/move the generated PDF to the desired file name and location
                Files.move(source, target);

                //System.out.println("PDF created successfully at: " + pdfFilePath);
                
            } else {
               // System.err.println("LibreOffice conversion failed with exit code: " + exitCode);
            }
            process.destroy();
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
    
    private static boolean isRangeAlreadyMerged(Sheet sheet, CellRangeAddress range) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == range.getFirstRow()
                    && mergedRegion.getLastRow() == range.getLastRow()
                    && mergedRegion.getFirstColumn() == range.getFirstColumn()
                    && mergedRegion.getLastColumn() == range.getLastColumn()) {
                return true; // The range is already merged
            }
        }
        return false; // The range is not merged
    }
    private static void setAllStyles(XSSFWorkbook workbook) {
    	blackColor = new XSSFColor(new Color(0, 0, 0), null);
		style = createCellStyle(workbook, true, HorizontalAlignment.CENTER, BorderStyle.THIN, blackColor, null, null);
		localHeadingStyle = createCellStyle(workbook, true, HorizontalAlignment.CENTER, null, null, null, null);
        greenColor = new XSSFColor(new Color(0, 128, 0), null);
        styleWithGreenBorder = createCellStyle(workbook, true, HorizontalAlignment.CENTER, BorderStyle.MEDIUM, greenColor, null, null);
        redColor = new XSSFColor(new Color(255, 0, 0), null);
        styleWithRedBorder = createCellStyle(workbook, true, HorizontalAlignment.CENTER, BorderStyle.MEDIUM, redColor, null, null);
        yellowColor = new XSSFColor(new Color(255, 255, 0), null);
        styleWithYellowBorder = createCellStyle(workbook, true, HorizontalAlignment.CENTER, BorderStyle.MEDIUM, yellowColor, null, null);
        mainheadingGreenColor = new XSSFColor(new Color(39, 78, 19), null);
        whiteColor = new XSSFColor(new Color(255, 255, 255), null);
        headingGreenColor = new XSSFColor(new Color(212, 228, 206), null);
        styleMainHeading = createCellStyle(workbook, true, HorizontalAlignment.CENTER, null, null, whiteColor, mainheadingGreenColor);
        styleHeading = createCellStyle(workbook, true, HorizontalAlignment.CENTER, null, null, blackColor, headingGreenColor);
        headingBlueColor = new XSSFColor(new Color(207, 226, 243), null);
        blueStyleHeading = createCellStyle(workbook, true, HorizontalAlignment.CENTER, null, null, blackColor, headingBlueColor);
        return;
    }
}
