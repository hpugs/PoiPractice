package com.hpugs.poi.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * @Description Poi测试
 * @author 高尚
 * @version 1.0
 * @date 创建时间：2018年1月26日 上午10:18:56
 */
public class Demo1 {
	
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
	 * @Description 测试创建Sheet页
	 * @throws IOException
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 上午11:08:03
	 */
	@Test
	public void createSheet() throws IOException{
		Workbook wb = new HSSFWorkbook();//定义一个新的工作簿
		
		//创建Sheet页
		wb.createSheet("第一个sheet页");
		wb.createSheet();
		wb.setSheetName(1, "第二个sheet页");
		wb.createSheet();
		wb.setSheetName(2, "第三个sheet页");
		wb.createSheet();
		wb.setSheetName(3, "第四个sheet页");
		wb.createSheet();
		wb.setSheetName(4, "第五个sheet页");
		
		//选中第二个Sheet页
		wb.setSelectedTab(1);
		
		//设置第二个Sheet页隐藏
		wb.setSheetHidden(3, true);
		
		//设置活动工作簿，表示是当前活动的工作表,一般是手工点击了某个工作表标签,则该工作表就成为了活动工作表了,也可以在程序中通过Activate方法使某个工作表成为活动工作表,这个集合只包含一个工作表
//		wb.setActiveSheet(4);
		
		System.out.println("当前工作簿有："+wb.getNumberOfSheets()+"个Sheet页");
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 测试创建单元格
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 上午11:14:23
	 */
	@Test
	public void createRow() throws IOException{
		Workbook wb = new HSSFWorkbook();//定义一个新的工作簿
		Sheet sheet = wb.createSheet("第一个Sheet页");
		
		//创建第一行Row
		Row row = sheet.createRow(0);
		//创建第一行第一列
		Cell cell = row.createCell(0);
		cell.setCellValue("第一行第一列");
		
		//创建第一行第二列
		row.createCell(1).setCellValue("第一行第二列");
		
		//创建第一行第三列
		row.createCell(2).setCellValue(true);
		
		//创建第一行第四列
		row.createCell(3).setCellValue(new Date());
		
		//创建第二行Row
		Row row1 = sheet.createRow(1);
		//创建第二行第一列
		Cell cell1 = row1.createCell(0);
		cell1.setCellValue(0);
		
		//创建第二行第二列
		row1.createCell(1).setCellValue(0.1);
		
		//创建第二行第三列
		row1.createCell(2).setCellValue(HSSFCell.CELL_TYPE_BLANK);
		
		wb.write(fileOut);
	}
	
	/**
	 * @Description 设置单元格时间格式
	 * @throws IOException
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 上午11:57:00
	 */
	@Test
	public void setRowValueDate() throws IOException{
		Workbook wb = new HSSFWorkbook();//定义一个新的工作簿
		Sheet sheet = wb.createSheet("第一个Sheet页");
		
		//创建第一行Row
		Row row = sheet.createRow(0);
		//创建第一行第一列
		Cell cell = row.createCell(0);
		cell.setCellValue(new Date());
		
		//创建时间格式
		CreationHelper creationHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();//创建单元格样式
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy年MM月dd日 HH:mm:ss"));
		
		//创建第一行第二列
		cell = row.createCell(1);
		cell.setCellValue(new Date());
		cell.setCellStyle(cellStyle);
		
		//创建第一行第三列
		cell = row.createCell(2);
		cell.setCellValue(Calendar.getInstance());
		cell.setCellStyle(cellStyle);
		
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
