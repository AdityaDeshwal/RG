package Report_Generation.RG;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.SwingWorker;
import javax.swing.JPanel;


public class App {

	private JFrame frame;
	private JTextField txtReportIn;
	public static JLabel lblProgress;
	private JButton btnRead;
    private JButton btnCreateReports;
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
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        int screenWidth = screenSize.width;
        int screenHeight = screenSize.height;

        // Calculate the frame dimensions for half the screen width and full height
        int frameWidth = screenWidth / 2;
        int frameHeight = screenHeight;
        int frameX = screenWidth / 2; // Start at the middle of the screen
        int frameY = 0; // Start at the top of the screen

        // Create and set up the frame
        frame = new JFrame();
        frame.setBounds(frameX, frameY, frameWidth, frameHeight);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(null);
		
		 JTextField txtAddressOfResult = new JTextField();
		 txtAddressOfResult.setToolTipText("");
	        txtAddressOfResult.setBounds(182, 212, 300, 25);
	        frame.getContentPane().add(txtAddressOfResult);
	        
	        JButton btnBrowseAddressOfResult = new JButton("Browse");
	        btnBrowseAddressOfResult.setBounds(564, 212, 100, 25);
	        frame.getContentPane().add(btnBrowseAddressOfResult);
		
		btnRead = new JButton("Read");
		btnRead.setBounds(285, 388, 85, 21);
		frame.getContentPane().add(btnRead);
		
//		txtReportIn = new JTextField();
//		txtReportIn.setBounds(182, 303, 294, 19);
//		frame.getContentPane().add(txtReportIn);
//		txtReportIn.setColumns(10);
		
//		JButton btnBrowseReportIn = new JButton("Browse");
//	    btnBrowseReportIn.setBounds(564, 303, 100, 19);
//	    frame.getContentPane().add(btnBrowseReportIn);
		
		btnCreateReports = new JButton("Create Reports");
		btnCreateReports.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		btnCreateReports.setBounds(225, 515, 197, 21);
		frame.getContentPane().add(btnCreateReports);
		
		
		lblProgress = new JLabel("...");
        lblProgress.setBounds(182, 350, 300, 25);
        frame.getContentPane().add(lblProgress);
		
		
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setPreferredSize(new Dimension(800, 600));
	    
	    btnBrowseAddressOfResult.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	            int result = fileChooser.showOpenDialog(frame);
	            if (result == JFileChooser.APPROVE_OPTION) {
	                File selectedFile = fileChooser.getSelectedFile();
	                txtAddressOfResult.setText(selectedFile.getAbsolutePath());
	            }
	        }
	    });
		
		btnRead.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
            	btnRead.setEnabled(false);
                new ReadWorker(txtAddressOfResult.getText()).execute();
            }
        });
		btnCreateReports.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
            	btnCreateReports.setEnabled(false);
                new CreateReportsWorker().execute();
            }
        });

	}
	public static void updateProgress(String text) {
        lblProgress.setText(text);
    }
	
    private class ReadWorker extends SwingWorker<Void, Void> {
        private String addressOfResult;
        //private String reportIn;

        public ReadWorker(String addressOfResult) {
            this.addressOfResult = addressOfResult;
        }

        @Override
        protected Void doInBackground() throws Exception {
            updateProgress("Reading and calculating results for students");
            String fileName = addressOfResult.replace("\"", "");
            SetBTestData.setBTestData(fileName);
            ConvertFormat.ConvertFormat(fileName);
            Generating_Report.adjusting_data(fileName);
//            updateProgress("Results are in memory."/* Please check the"
//            		+ " sample report in excel file and do changes if required."
//            		+ " Please make sure to SAVE after changes. "
//            		+ "After all this, please click Create Reports button.*/);
            return null;
        }

        @Override
        protected void done() {
            //btnRead.setEnabled(true);
        	updateProgress("Results are in memory."/* Please check the"
            		+ " sample report in excel file and do changes if required."
            		+ " Please make sure to SAVE after changes. "
            		+ "After all this, please click Create Reports button.*/);
        }
    }

    private class CreateReportsWorker extends SwingWorker<Void, Void> {

        @Override
        protected Void doInBackground() throws Exception {
            Generating_Report.createReportsPdf();
            return null;
        }

        @Override
        protected void done() {
        	updateProgress("Reports Completed");
        	btnCreateReports.setEnabled(true);
        }
    }
}
