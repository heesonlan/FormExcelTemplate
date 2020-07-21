package com.lan.excel.main.report;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.lan.excel.util.FormExcelUtil;

/**
 * 
 * @author LAN
 * @date 2020年7月20日
 */
public class FormTemplateExcel {
	Logger log = LoggerFactory.getLogger(getClass());

	public static void main(String[] args) {
		System.out.println("start...");
		new FormTemplateExcel().generateRiskReport();
		System.out.println("finish");
	}
	
	private void generateRiskReport(){
		InputStream in = null;
		XSSFWorkbook wb = null;
		try {
			/****读取Excel模板*******************************/
			in =Thread.currentThread().getContextClassLoader().getResourceAsStream("excel/模板.xlsx");
			wb = new XSSFWorkbook(in);
			/****填入数据************************************/
			genarateReport(wb);
			/****另存为新文件*********************************/
			saveExcelToDisk(wb, "D:\\data\\excel\\报告.xlsx");
			
		} catch (IOException e) {
			log.error("error", e);
		} finally {
			try {if(wb!=null)wb.close();} catch (IOException e) { log.error("error", e);}
			try {if(in!=null)in.close();} catch (IOException e) { log.error("error", e);}
		}
	}
	
	private void genarateReport(XSSFWorkbook wb) {
		XSSFSheet sheet1 = wb.getSheetAt(0);
		XSSFSheet sheet2 = wb.getSheetAt(1);
		// 设置公式自动读取
		sheet1.setForceFormulaRecalculation(true);
		sheet2.setForceFormulaRecalculation(true);
		
		/***设置单个单元格内容*********************************/
		FormExcelUtil.setCellData(sheet1, "2020-07报告", 1, 1);
		/***第一个表格*********************************/
		ExampleData ea = new ExampleData();
		List<List<Object>> data1 = ea.getData1(10);
		int addRows=0;
		//动态插入行
		//FormExcelUtil.insertRowsStyleBatch(sheet, startNum, insertRows, styleRow, styleColStart, styleColEnd)
		//按照styleRow行的格式，在startNum行后添加insertRows行，并且针对styleColStart~ styleColEnd列同步模板行styleRow的格式
		FormExcelUtil.insertRowsStyleBatch(sheet1, 4+addRows, data1.size()-2, 4, 1, 4);
		FormExcelUtil.setTableData(sheet1, data1, 4+addRows, 1);
		addRows += data1.size()-2;
		/***第二个表格*********************************/
		List<List<Object>> data2 = ea.getData2();
		FormExcelUtil.insertRowsStyleBatch(sheet1, 10+addRows, data2.size()-2, 10+addRows, 1, 6);
		FormExcelUtil.setTableData(sheet1, data2, 10+addRows, 1);
		addRows += data2.size()-2;
		/***第三个表格*********************************/
		List<List<Object>> data3 = ea.getData3();
		FormExcelUtil.setTableData(sheet2, data3, 3, 1);
		
	}

	private void saveExcelToDisk(XSSFWorkbook wb, String filePath){
		File file = new File(filePath);
		OutputStream os=null;
		try {
			os = new FileOutputStream(file);
			wb.write(os);
			os.flush();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {if(os!=null)os.close();} catch (IOException e) { log.error("error", e);}
		}
	}
}
