package com.lan.excel.util;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.ResourceUtils;
import org.springframework.util.StringUtils;

public class FormExcelUtil {
	//
	public static void setCellData(XSSFSheet sheet, Object data, int rowNo, int colNo) {

		try {

			// 向相应的单元格里面设置值
			XSSFRow row;
			row = sheet.getRow(rowNo-1);
			XSSFCell cell = row.getCell(colNo - 1);
			if(cell == null){
				cell = row.createCell(colNo - 1);
			}
			if(data == null){

			}else if(data instanceof String){
				cell.setCellValue(data.toString());
			}else if(data instanceof Integer){
				cell.setCellValue((Integer)data);
			}else if(data instanceof Float){
				cell.setCellValue((Float)data);
			}else if(data instanceof Double){
				cell.setCellValue((Double)data);
			}else if(data instanceof BigDecimal){
				cell.setCellValue(Double.valueOf(data.toString()));
			}else{
				throw new RuntimeException("传入excel表格中的数据只能是String，Integer，Float，Double，BigDecimal类型，当前传入了"+data.getClass().getSimpleName());
			}
		} catch (Exception e){
			throw e;
		}
	}
	//
	public static void setCellStyle(XSSFSheet sheet, XSSFCellStyle style, int rowNo, int colNo) {
		XSSFRow row;
		row = sheet.getRow(rowNo - 1);
		XSSFCell cell = row.getCell(colNo - 1);
		if(cell == null){
			cell = row.createCell(colNo - 1);
		}
		cell.setCellStyle(style);
	}
	//
	public static XSSFCellStyle getCellStyle(XSSFSheet sheet, int rowNo, int colNo) {
		XSSFRow row;
		row = sheet.getRow(rowNo - 1);
		XSSFCell cell = row.getCell(colNo - 1);
		if(cell == null){
			cell = row.createCell(colNo - 1);
		}
		return cell.getCellStyle();
	}
	//
	public static void setCellFormula(XSSFSheet sheet, String formula, int rowNo, int colNo) {
		if(formula == null || "".equals(formula)) return;
		XSSFRow row;
		row = sheet.getRow(rowNo - 1);
		XSSFCell cell = row.getCell(colNo - 1);
		if(cell == null){
			cell = row.createCell(colNo - 1);
		}
		//cell.setCellType(CellType.FORMULA);
		cell.setCellFormula(formula);
	}
	//
	public static void setTableData(XSSFSheet sheet, List<List<Object>> tableData, int fromRow, int fromCol) {
		// 向相应的单元格里面设置值
		XSSFRow row;
		int orgCol = fromCol;
		for(List<Object> l1 : tableData){
			fromCol = orgCol;
			row = sheet.getRow(fromRow-1);
			for(Object o:l1){
				if(o==null){
					//do nothing
				}else if(o instanceof String){
					row.getCell(fromCol-1).setCellValue(o.toString());
				}else if(o instanceof Integer){
					row.getCell(fromCol-1).setCellValue((Integer)o);
				}else if(o instanceof Float){
					row.getCell(fromCol-1).setCellValue((Float)o);
				}else if(o instanceof Double){
					row.getCell(fromCol-1).setCellValue((Double)o);
				}else if(o instanceof BigDecimal){
					row.getCell(fromCol-1).setCellValue(Double.valueOf(o.toString()));
				}else if(o instanceof Long){
					row.getCell(fromCol-1).setCellValue((Long)o);
				}else{
					throw new RuntimeException("传入excel表格中的数据只能是String，Integer，Float，Double，BigDecimal类型，当前传入了"+o.getClass().getSimpleName());
				}
				fromCol++;
			}
			fromRow++;
		}		
	}
	//
	public static void setRowData(XSSFSheet sheet, List<Object> rowData, int rowIndex, int fromCol) {
		// 向相应的单元格里面设置值
		XSSFRow row = sheet.getRow(rowIndex-1);
		for(Object o:rowData){
			if(o==null){
				
			}else if(o instanceof String){
				row.getCell(fromCol-1).setCellValue(o.toString());
			}else if(o instanceof Integer){
				row.getCell(fromCol-1).setCellValue((Integer)o);
			}else if(o instanceof Float){
				row.getCell(fromCol-1).setCellValue((Float)o);
			}else if(o instanceof Double){
				row.getCell(fromCol-1).setCellValue((Double)o);
			}else if(o instanceof BigDecimal){
				row.getCell(fromCol-1).setCellValue(Double.valueOf(o.toString()));
			}else if(o instanceof Long){
				row.getCell(fromCol-1).setCellValue((Long)o);
			}else{
				throw new RuntimeException("传入excel表格中的数据只能是String，Integer，Float，Double，BigDecimal类型，当前传入了"+o.getClass().getSimpleName());
			}
			fromCol++;
		}			
	}
	//
	public static void insertImg(XSSFWorkbook wb, Drawing drawing, String imgPath, int fromRow, int fromCol) {
		if(StringUtils.isEmpty(imgPath)){
			return;
		}
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();

		try {
			
			BufferedImage bufferImg = ImageIO.read(ResourceUtils.getFile(imgPath));
		
	        //构建图片流
	        //BufferedImage tag = new BufferedImage(320, 240, BufferedImage.TYPE_INT_RGB);
	        //绘制改变尺寸后的图
	        //tag.getGraphics().drawImage(bufferImg, 0, 0,320, 240, null);
	
	        ImageIO.write(bufferImg, "png", byteArrayOut);
	
	        int pictureIdx = wb.addPicture(byteArrayOut.toByteArray(), wb.PICTURE_TYPE_PNG);
	        
	        CreationHelper helper = wb.getCreationHelper();
	        ClientAnchor anchor = helper.createClientAnchor();
	        anchor.setRow1(fromRow-1);// 高度
	        anchor.setCol1(fromCol-1);//距离
	        Picture pict = drawing.createPicture(anchor, pictureIdx);
	        pict.resize();

		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	//
	public static void deleteRow(XSSFSheet sheet, int startRow, int endRow){

		if(startRow > endRow) return;
		int start = startRow - 1;
		int lastRowNum = sheet.getLastRowNum();

		for(int i = start; i <= (start + (endRow - startRow)); i++){
			if(i > lastRowNum)break;
			XSSFRow row = sheet.getRow(i);
			if(row == null)continue;
			if(!isRowEmpty(row)){
				//此方法只能删除行内容
				sheet.removeRow(row);
			}
		}
		if(endRow <= lastRowNum){
			//此方法是将余下的行向上移
			sheet.shiftRows(endRow, lastRowNum, (startRow - endRow - 1));
		}

	}
	//
	public static boolean isRowEmpty(XSSFRow row) {

		for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
			XSSFCell cell = row.getCell(c);
			if (cell != null && cell.getCellType() != CellType.BLANK)
				return false;
		}
		return true;
	}
	//
	public static void insertRowsBatch(XSSFSheet sheet, int startNum, int endNum) {
		if(startNum < 0 || endNum < 0  || startNum > endNum) return;

		if (sheet.getRow(startNum) != null) {
			int lastRowNo = sheet.getLastRowNum();
			sheet.shiftRows(startNum, lastRowNo, (endNum - startNum));
		}
		//
		for(int i = startNum; i < endNum; i++){
			createRow(sheet, i);
		}
	}
	//
	private static XSSFRow createRow(XSSFSheet sheet, Integer rowIndex) {
		XSSFRow row = null;
		row = sheet.createRow(rowIndex);
		return row;
	}

	/**
	 * 按照styleRow模板行的格式，在startNum行后添加insertRows行，并且针对styleColStart~ styleColEnd列同步模板行styleRow的格式
	 * @author LAN
	 * @date 2020年3月20日
	 * @param sheet
	 * @param startNum
	 * @param insertRows
	 * @param styleRow
	 * @param styleColStart
	 * @param styleColEnd
	 */
	public static void insertRowsStyleBatch(XSSFSheet sheet, int startNum, int insertRows, int styleRow, int styleColStart, int styleColEnd) {
		int endNum = startNum+insertRows;
		if(startNum < 0 || endNum < 0  || startNum > endNum) return;
		insertRowsBatch(sheet, startNum, endNum);
		
		for(int i=styleColStart; i<=styleColEnd; i++) {
			XSSFCellStyle cellStyle = getCellStyle(sheet, styleRow, i);
			for(int j=startNum; j<=endNum; j++) {
				setCellStyle(sheet, cellStyle, j, i);
			}
		}
	}
}
