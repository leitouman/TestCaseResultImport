package com.wx.ui;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Timer;
import java.util.TimerTask;
import java.util.UUID;

import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.mks.api.response.APIException;
import com.wx.service.ExcelUtil;
import com.wx.service.MyRunnable;
import com.wx.util.APIExceptionUtil;
import com.wx.util.Constants;
import com.wx.util.MKSCommand;

import javax.swing.SwingConstants;

public class TestCaseImport extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	public static JPanel contentPane;
	private JTabbedPane tabbedPane;
	private JTable tableMapper;
	private JButton backBtn;
	private JButton nextBtn;
	private JButton doneBtn;
	private static JTextArea textArea;
	private JLabel pathText;
	private static MKSCommand cmd;
	private static final Map<String, String> ENVIRONMENTVAR = System.getenv();
	public static final Logger logger = Logger.getLogger(TestCaseImport.class.getName());
	private static String defaultUser = "admin"; 
	private static JLabel helloText;
	public String documentTitle = null;// 用来存放文档标题
	static String project =null;
	private File excelFile;
	private JTextField testSuiteField;
	private String testSuiteID;
	private List<List<Map<String, Object>>> datas;
	private List<List<Map<String, Object>>> realData ;
	private ExcelUtil excelUtil = new ExcelUtil();
	//public static String TOKEN ;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());// 界面风格
					TestCaseImport frame = new TestCaseImport();
					frame.setVisible(true);
					frame.setLocationRelativeTo(null);
					
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * 初始化MKSCommand中的参数
	 */
	public void setMksConfig() {
		try {
			String host = TestCaseImport.ENVIRONMENTVAR.get(Constants.MKSSI_HOST);
			if(host==null || host.length()==0) {
				host = "172.25.4.18";
			}
			String portStr = ENVIRONMENTVAR.get(Constants.MKSSI_PORT);
			Integer port = portStr!=null && !"".equals(portStr)? Integer.valueOf(portStr) : 7001;
			defaultUser = ENVIRONMENTVAR.get(Constants.MKSSI_USER);
			String pwd = "";
			if(defaultUser == null || "".equals(defaultUser) ){
				defaultUser = "admin";
				pwd = "admin";
			}
			cmd = new MKSCommand(host, port, defaultUser, pwd, 4, 16);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestCaseImport.contentPane, "Can not get a connection!", "Message",
					JOptionPane.WARNING_MESSAGE);
			TestCaseImport.logger.info("Can not get a connection!");
			System.exit(0);
		}
	}

//	/**
//	 * 获得连接
//	 * 
//	 * @param c
//	 */
//	public boolean getConnect() {
//		try {
//			
//			this.setMksConfig();
//			
//		} catch (Exception e) {
//			JOptionPane.showMessageDialog(this, "Can not get a connection!", "Message", JOptionPane.WARNING_MESSAGE);
//			return false;
//		}
//		return true;
//	}

	/**
	 * Create the frame.
	 * @throws Exception 
	 */
	public TestCaseImport() throws Exception {
		
		setTitle("Excel Import");
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 849, 416);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane.setBounds(8, 8, 834, 322);
		contentPane.add(tabbedPane);
		
		
		JPanel panel = new JPanel();
		panel.setForeground(Color.RED);
		panel.setToolTipText("Test Suite");
		tabbedPane.addTab(" Info ", null, panel, null);
		panel.setLayout(null);
		
		pathText = new JLabel("<Path to Excel File *.xls>");
		pathText.setBounds(25, 227, 648, 24);
		pathText.setBorder(BorderFactory.createEtchedBorder());
		panel.add(pathText);
		
		JButton browseBtn = new JButton("Browse");
		browseBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				logger.info("Start to load excel");
				helloText.setText("Loading Excel File");
				helloText.setForeground(Color.BLACK);
				JFileChooser fc = new JFileChooser();
				fc.setDialogTitle("Select Excel File");
				fc.setAcceptAllFileFilterUsed(true);
				fc.setMultiSelectionEnabled(false);
				int returnVal = fc.showOpenDialog(contentPane);
				if (returnVal == 0) {
					excelFile = fc.getSelectedFile();
					String path = excelFile.getAbsolutePath();
					if (!path.endsWith("xls") && !path.endsWith("xlsx")) {
						logger.error("Selected file is not a excel file!");
						JOptionPane.showMessageDialog(contentPane, "Please Choose Excel File",
								"Please Choose Excel File", JOptionPane.ERROR_MESSAGE);
						helloText.setText("Please Choose Excel File!");
						helloText.setForeground(Color.RED);
					} else {
						String suiteId = path.substring(path.lastIndexOf("-") + 1, path.lastIndexOf("."));
						pathText.setText(path);
						if (suiteId.matches("[0-9]*")) {
							testSuiteID = suiteId;
							testSuiteField.setText(suiteId);
						}
						try {
							datas = excelUtil.parseExcel(excelFile);
						} catch (Exception e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
					}
				}
			}
		});
		browseBtn.setBounds(696, 226, 89, 27);
		panel.add(browseBtn);
		
		JLabel lblTestSuite = new JLabel("Test Suite ID : ");
		lblTestSuite.setBounds(25, 50, 144, 24);
		panel.add(lblTestSuite);
		
		testSuiteField = new JTextField();
		testSuiteField.setBounds(188, 49, 485, 27);
		panel.add(testSuiteField);
		testSuiteField.setColumns(10);
		
		JLabel lblNewLabel = new JLabel("Project       :");
		lblNewLabel.setBounds(25, 116, 144, 24);
		panel.add(lblNewLabel);
		
		lblTheProject = new JLabel("( The project must be fill in while importing new document. Format Example: /Test1 )");
		lblTheProject.setVerticalAlignment(SwingConstants.TOP);
		lblTheProject.setForeground(Color.RED);
		lblTheProject.setBounds(25, 170, 822, 24);
		panel.add(lblTheProject);
		
		comboBox_2 = new JComboBox();
		
		comboBox_2.setBounds(188, 115, 485, 27);
		panel.add(comboBox_2);
		
		JLabel lblNewLabel_1 = new JLabel("*");
		lblNewLabel_1.setForeground(Color.BLUE);
		lblNewLabel_1.setBounds(168, 52, 18, 21);
		panel.add(lblNewLabel_1);
		
		JLabel label = new JLabel("*");
		label.setForeground(Color.BLUE);
		label.setBounds(168, 118, 18, 21);
		panel.add(label);
		Object obj = comboBox_2.getSelectedItem();
		JPanel panel_1 = new JPanel();
		tabbedPane.addTab(" Mapping ", null, panel_1, null);
		panel_1.setLayout(null);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(0, 0, 975, 378);
		panel_1.add(scrollPane);

		tableMapper = new JTable();
		tableMapper.setModel(
				new DefaultTableModel(new Object[][] { new Object[2], new Object[2], new Object[2], new Object[2] },
						new String[] { "Excel Headers", "Integrity Fields" }));
		scrollPane.setViewportView(tableMapper);

		JPanel panel_2 = new JPanel();
		tabbedPane.addTab(" Logger ", null, panel_2, null);
		panel_2.setLayout(null);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setBounds(-1, -1, 977, 385);
		panel_2.add(scrollPane_1);

		textArea = new JTextArea();
		textArea.setLineWrap(true);
		scrollPane_1.setViewportView(textArea);
		project = obj!=null ? obj.toString() : "";

		doneBtn = new JButton("Done");
		doneBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});
		doneBtn.setEnabled(false);
		doneBtn.setBounds(677, 345, 100, 27);
		contentPane.add(doneBtn);
		
		nextBtn = new JButton("Next");
		nextBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				nextAction(1);
			}
		});
		nextBtn.setBounds(558, 345, 100, 27);
		contentPane.add(nextBtn);

		backBtn = new JButton("back");
		backBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				nextAction(-1);
			}
		});
		backBtn.setEnabled(false);
		backBtn.setBounds(428, 345, 100, 27);
		contentPane.add(backBtn);

		helloText = new JLabel("Hello :)");
		helloText.setBounds(31, 345, 337, 18);
		contentPane.add(helloText);

		
		try {
			setMksConfig();
			
			setProjectList();
		} catch (Exception e1) {
			JOptionPane.showMessageDialog(this, e1.getMessage());
			e1.printStackTrace();
		}
		for (int i = 0; i < tabbedPane.getMouseListeners().length; i++) {
			tabbedPane.removeMouseListener(tabbedPane.getMouseListeners()[i]);
		}
		excelUtil.parsFieldMapping();
	}

	public void nextAction(int plus) {
		int curIdx = tabbedPane.getSelectedIndex();
		int maxIdx = tabbedPane.getComponentCount() - 1;
		int newIdx = curIdx + plus;
		boolean pass = true;
		if (newIdx == 1) {
			
			// 进入选择界面，需要判断MKS是否输入
			logger.info("==> into Mapper panel");
			// 检查excel是否解析成功
			if (excelFile == null) {
				JOptionPane.showConfirmDialog(contentPane, "Please select a excel!");
				return;
			}

			// 判断是否选择模板类型
			// 解析Excel
			testSuiteID = testSuiteField.getText();
			/*
			 * if(testSuiteID == null || "".equals(testSuiteID)) {
			 * JOptionPane.showConfirmDialog(contentPane,
			 * "Test Suite ID can't Be Empty!"); return; }
			 */
			
			if (testSuiteID == null || "".equals(testSuiteID)) {
				if ( documentTitle == null || "".equals(documentTitle)) {
					documentTitle = JOptionPane.showInputDialog(
							"Document ID Is Empty, So Please Enter [ Document Short Title ] to Create It!", documentTitle);
					if (documentTitle == null || documentTitle.equals("")) {
						JOptionPane.showInputDialog(
								"Document ID and  [ Document Short Title ] Counld Not Be Empty Simultaneously");
					}		
				}
				project=comboBox_2.getSelectedItem().toString();
				if (project == "Please select a Project" || "Please select a Project".equals(project)) {
					JOptionPane.showMessageDialog(this, "Please select a Project!");
					comboBox_2.addActionListener(new ActionListener() {
						public void actionPerformed(ActionEvent arg0) {
							project=comboBox_2.getSelectedItem().toString();
						}
					});
					return;
				} else {
					boolean projectHas = false;
					try {
						projectHas = cmd.checkProject(project);
					} catch (APIException e) {
						logger.info(e.getMessage());
					}
					if (!projectHas) {
						JOptionPane.showMessageDialog(this, "Project is not exist, Please Re-Input It!");
						return;
					}
				}
			} else {
				try {
					if (!cmd.docIDIsRight( testSuiteID, "Test Suite") ) {// 此处要修改，  判断类型
						JOptionPane.showConfirmDialog(contentPane,
								"Your input Test Suite ID is not correctly, Please Re-Input It!");
						return;
					}
				} catch (Exception e1) {
					JOptionPane.showConfirmDialog(contentPane,
							"Your input Test Suite ID is not correctly, Please Re-Input It!");
					return;
				}
			}

			if (datas == null || datas.size() == 0) {
				JOptionPane.showConfirmDialog(contentPane, "Counld not prase excel! Please check the excel format!");
				return;
			}
			List<List<Map<String, Object>>> dealDatas = excelUtil.dealExcelData(datas);
			try {
				tableMapper.setModel(new DefaultTableModel(excelUtil.tableFields,
						new String[] { "Excel Headers", "Integrity Fields" }));
				Map<String,String> errorRecord = new HashMap<String,String>();
				realData = excelUtil.checkExcelData(dealDatas, errorRecord, cmd);
				String checkMessage = errorRecord.get("error");
				if(checkMessage != null && !"".equals(checkMessage)){
					JOptionPane.showMessageDialog(this, checkMessage);
					return;
				}
			} catch (APIException e) {
				APIExceptionUtil.getMsg(e);
			} catch (Exception e){
				e.printStackTrace();
			}
		}
		if (newIdx == 2) {
			// 进入Logger界面
			logger.info("==> into logger panel");
			// 开始线程导入数据
			r.cmd = cmd;
			r.datas = realData;
			r.testSuiteId = testSuiteID;
			r.excelUtil = excelUtil;
			r.project = comboBox_2.getSelectedItem().toString();
			r.shortTitle = documentTitle;
			t = new Thread(r);
			t.start();// t查询线程,开启
			configTimeArea(j);
		}
		if (pass) {
			if (newIdx > 0) {
				backBtn.setEnabled(true);
			} else {
				backBtn.setEnabled(false);
			}
			if (newIdx < maxIdx) {
				nextBtn.setEnabled(true);
				doneBtn.setEnabled(false);
			}
			if (newIdx == maxIdx) {
				nextBtn.setEnabled(false);
			}
			if (newIdx <= maxIdx) {
				setFocus(newIdx);
			}
		}
	}

	/**
	 * 选择tab
	 * 
	 * @param idx
	 */
	private void setFocus(final int idx) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				tabbedPane.setSelectedIndex(idx);
			}
		});
	}

	/**
	 * 选择tab
	 * 
	 * @param idx
	 */
	public static void showLogger(final String logger) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				textArea.append(logger + "\n");
			}
		});
	}

	/**
	 * 选择tab
	 * 
	 * @param idx
	 */
	public static void showProgress(final int sheetNum, final int totalSheetCount, final int caseNum,
			final int totalCaseCount) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				helloText.setText("Task : " + sheetNum + "/" + totalSheetCount + ", Progress : " + caseNum + "/"
						+ totalCaseCount);
			}
		});
	}

	private int ONE_SECOND = 1000;
	private MyRunnable r = new MyRunnable();
	private Thread t = new Thread();// 查询线程
	private JLabelTimerTask j = new JLabelTimerTask();
	private JLabel lblTheProject;
	private static JComboBox comboBox_2;

	/**
	 * 这个方法创建 a timer task 每秒更新一次 the time
	 */
	private void configTimeArea(JLabelTimerTask j) {
		Timer tmr = new Timer();
		tmr.scheduleAtFixedRate(j, new Date(), ONE_SECOND);
	}

	/**
	 * Timer task 更新时间显示区
	 * 
	 */
	protected class JLabelTimerTask extends TimerTask {
		@Override
		public void run() {
			if (!t.isAlive()) {
				doneBtn.setEnabled(true);
			} else {
				doneBtn.setEnabled(false);
			}
		}
	}

	
	@SuppressWarnings("unchecked")
	public static void setProjectList() throws APIException{
		List<String> projects = cmd.getProjects(defaultUser);
		projects.add(0, "Please select a Project");
		comboBox_2.setModel(new DefaultComboBoxModel<String>(projects.toArray(new String[projects.size()])));
	}
}
