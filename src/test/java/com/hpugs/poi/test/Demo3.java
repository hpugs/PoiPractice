package com.hpugs.poi.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * @Description 单元格样式调整
 * @author 高尚
 * @version 1.0
 * @date 创建时间：2018年1月26日 下午1:40:48
 */
public class Demo3 {
	
	private static final String filePath = "F:\\poi\\poi测试.xls";
	private static FileOutputStream fileOut;
	
	@Before
	public void beforeTest() throws IOException{
		File file = new File(filePath);
		if(file.exists()){
			file.delete();
			file.createNewFile();
		}
		fileOut = new FileOutputStream(file);
	}
	
	/**
	 * @Description 测试创建工作簿
	 * @throws IOException
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 上午10:47:53
	 */
	@Test
	public void createWorkbook() throws IOException{
		Workbook wb = new HSSFWorkbook();//定义一个新的工作簿
		wb.write(fileOut);
	}
	
	/**
	 * @Description 测试设置行高
	 * @throws IOException
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 上午11:08:03
	 */
	@Test
	public void setRowHeight() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();//定义一个新的工作簿
		HSSFSheet sheet = wb.createSheet("第一个sheet页");
		
		HSSFRow row = sheet.createRow(1);
		row.setHeightInPoints(50);
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 测试设置单元格的样式
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 下午1:58:31
	 */
	@Test
	public void setCellStyle() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();//定义一个新的工作簿
		HSSFSheet sheet = wb.createSheet("第一个sheet页");
		
		HSSFRow row = sheet.createRow(1);
		row.setHeightInPoints(50);
		Cell cell = createCellStyle(wb, row, 1, HSSFCellStyle.ALIGN_CENTER, HSSFCellStyle.VERTICAL_CENTER);
		cell.setCellValue("hpugs");
		
		cell = createCellStyle(wb, row, 2, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_BOTTOM);
		cell.setCellValue("hpugs");
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 创建一个单元格并为其设置对齐方式
	 * @param hssf 工作簿
	 * @param row 单元格行
	 * @param column 单元格index
	 * @param halign 水平对齐方式
	 * @param valign 垂直对齐方式
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 下午2:00:48
	 */
	private Cell createCellStyle(HSSFWorkbook hssf, HSSFRow row, int column, short halign, short valign){
		Cell cell = row.createCell(column);
		CellStyle cellStyle = hssf.createCellStyle();
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(valign);
		cell.setCellStyle(cellStyle);
		return cell;
	}
	
	/**
	 * @Description 设置单元格的borderStyle
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 下午2:14:55
	 */
	@Test
	public void setCellBorderStyle() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();//定义一个新的工作簿
		HSSFSheet sheet = wb.createSheet("第一个sheet页");
		
		HSSFRow row = sheet.createRow(1);
		row.setHeightInPoints(50);
		
		HSSFCell cell = row.createCell(1);
		
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		
		cellStyle.setBorderBottom(CellStyle.BORDER_DASHED);
		cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		
		cell.setCellStyle(cellStyle);
		cell.setCellValue("hpugs");
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 设置单元格背景色
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 下午2:25:00
	 */
	@Test
	public void setCellBackground() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();//定义一个新的工作簿
		HSSFSheet sheet = wb.createSheet("第一个sheet页");
		
		HSSFRow row = sheet.createRow(1);
		row.setHeightInPoints(50);
		
		HSSFCell cell = row.createCell(1);
		cell.setCellValue("hpugs");
		
		CellStyle cellStyle = wb.createCellStyle();
		//设置背景色
		cellStyle.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
		//设置前景色
		cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		//设置浮雕样式
		cellStyle.setFillPattern(CellStyle.ALIGN_CENTER);
		cell.setCellStyle(cellStyle);
		
		wb.write(fileOut);
	}
	
	/**
	 * 设置单元格的字体
	 * @throws IOException 
	 */
	@Test
	public void setTextStyle() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet shoot = wb.createSheet("第一个shoot页");
		HSSFRow row = shoot.createRow(0);
		HSSFCell cell = row.createCell(1);
		cell.setCellValue("今天是星期天");
		
		//设置字体样式
		Font font = wb.createFont();
//		font.setFontHeight((short)100);//设置字体高度
		font.setFontHeightInPoints((short)20);//设置字体以及单元格的高度
		font.setFontName("myAuto");
		font.setItalic(true);
		font.setStrikeout(true);
		
		//添加字体样式
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setFont(font);
		
		//设置单元格样式
		cell.setCellStyle(cellStyle);
		
		wb.write(fileOut);
	}
	
	/**
	 * 设置单元格自动换行
	 * @throws IOException 
	 */
	@Test
	public void setCellWrap() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet shoot = wb.createSheet("第一个shoot页");

		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setWrapText(true);//设置单元格换行
		
		HSSFRow row = shoot.createRow(1);
		HSSFCell cell = row.createCell(1);
		cell.setCellValue("今天是星期天真舒服");
		cell.setCellStyle(cellStyle);
		
		row = shoot.createRow(2);
		cell = row.createCell(1);
		cell.setCellValue("今天是星期天\n真舒服");
		cell.setCellStyle(cellStyle);
		
		wb.write(fileOut);
	}
	
	/**
	 * 设置单元格自动换行
	 * @throws IOException 
	 */
	@Test
	public void setDataFormat() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet shoot = wb.createSheet("第一个shoot页");

		DataFormat format = wb.createDataFormat();
		
		HSSFRow row = shoot.createRow(1);
		HSSFCell cell = row.createCell(1);
		cell.setCellValue(123456.25);
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(format.getFormat("0.0"));
		cell.setCellStyle(cellStyle);
		
		row = shoot.createRow(2);
		cell = row.createCell(1);
		cell.setCellValue(123456.25);
		cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(format.getFormat("#,##0.000"));
		cell.setCellStyle(cellStyle);
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 单元格合并
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 下午2:36:05
	 */
	@Test
	public void cellMerged() throws IOException{
		HSSFWorkbook wb = new HSSFWorkbook();//定义一个新的工作簿
		HSSFSheet sheet = wb.createSheet("第一个sheet页");
		
		HSSFRow row = sheet.createRow(1);
		row.setHeightInPoints(50);
		
		HSSFCell cell = row.createCell(1);
		cell.setCellValue("hpugs");
		
		sheet.addMergedRegion(new CellRangeAddress(
					1, //起始行
					1, //结束行
					1, //起始列
					2  //结束列
				));
		
		sheet.addMergedRegion(new CellRangeAddress(
				2, //起始行
				3, //结束行
				1, //起始列
				1  //结束列
			));
		
		sheet.addMergedRegion(new CellRangeAddress(
				5, //起始行
				6, //结束行
				1, //起始列
				2  //结束列
			));
		
		wb.write(fileOut);
	}
	
	@After
	public void afterTest() throws IOException{
		if(null != fileOut){
			fileOut.flush();
			fileOut.close();
		}
	}

}
