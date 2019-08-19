package com.wx.service;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import com.wx.ui.TestCaseImport;
import com.wx.util.MKSCommand;

public class MyRunnable implements Runnable {
	public MKSCommand cmd;
	public ExcelUtil excelUtil;
	public String importType;
	public String testSuiteId;
	public List<List<Map<String,Object>>> datas;
	public String project;
	public String shortTitle;
	public MyRunnable() {
		super();
	}

	@Override
	public void run() {
		try {
			TestCaseImport.logger.info("===============Start to import Test Case==============");
		//	if( TestCaseImport.TOKEN != null ) {
				excelUtil.startImport(datas, cmd, importType,shortTitle,project, testSuiteId);
		//	}
			JOptionPane.showMessageDialog(TestCaseImport.contentPane, "Done", "Success", JOptionPane.INFORMATION_MESSAGE);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(TestCaseImport.contentPane, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} finally {
			try {
				cmd.release();
			} catch (IOException e) {
				
			}
			TestCaseImport.logger.info("===============End to import Test Case==============");
		}
	}

	

}
