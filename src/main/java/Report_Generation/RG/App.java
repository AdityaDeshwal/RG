package Report_Generation.RG;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JTextField;


public class App {

	private JFrame frame;
	private JTextField txtReportIn;
	public static void main(String[] args) {
		//SpringApplication.run(App.class, args);
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					App window = new App();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public App() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 597, 395);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		 JTextField txtAddressOfResult = new JTextField();
		 txtAddressOfResult.setToolTipText("");
	        txtAddressOfResult.setBounds(128, 52, 300, 25);
	        frame.getContentPane().add(txtAddressOfResult);
		
		JButton btnRead = new JButton("Read");
		btnRead.setBounds(233, 167, 85, 21);
		frame.getContentPane().add(btnRead);
		
		txtReportIn = new JTextField();
		txtReportIn.setBounds(134, 107, 294, 19);
		frame.getContentPane().add(txtReportIn);
		txtReportIn.setColumns(10);
		
		JButton btnCreateReports = new JButton("Create Reports");
		btnCreateReports.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		btnCreateReports.setBounds(173, 260, 197, 21);
		frame.getContentPane().add(btnCreateReports);
		
		btnRead.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                //String fileName = txtAddressOfResult.getText().replace("\"", "");
            	btnRead.setEnabled(false);
                String fileName="C:\\Users\\adity\\Downloads\\For Aditya sir (1).xlsx".replace("\"","");
            	System.out.println("File path: " + fileName); // Debugging statement
                //reading_excel reader = new reading_excel();
                //reader.read(fileName);
                SetBTestData.setBTestData(fileName);
                System.out.println(SetBTestData.BTestData);
                ConvertFormat.ConvertFormat(fileName);
                
                //String fileNameReportIn = txtReportIn.getText().replace("\"", "");
                String fileNameReportIn="\"C:\\Users\\adity\\Downloads\\Copy of PCM Logic and Coding..xlsx\"".replace("\"","");
                Generating_Report.adjusting_data(fileNameReportIn);
            }
        });
		btnCreateReports.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
					Generating_Report.createReports();
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
            }
        });
		//"C:\Users\adity\Downloads\For Aditya sir.xlsx"
		//C:\Users\adity\Downloads\For Aditya sir.xlsx

	}
}
