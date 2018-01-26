package com.hpugs.poi.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * @Description 读取Excel
 * @author 高尚
 * @version 1.0
 * @date 创建时间：2018年1月26日 下午12:52:25
 */
public class Demo2 {
	
	private static final String filePath = "F:\\poi\\poi测试.xls";
	private static FileInputStream fileIn;
	
	@Before
	public void beforeTest() throws IOException{
		File file = new File(filePath);
		if(file.exists()){
			fileIn = new FileInputStream(filePath);
		}
	}
	
	/**
	 * @Description 获取当前Excel共有多少个Sheet页
	 * @throws Exception
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 下午12:58:38
	 */
	@Test
	public void readSheetCount() throws Exception{
		POIFSFileSystem fileSystem = new POIFSFileSystem(fileIn);
		HSSFWorkbook hssf = new HSSFWorkbook(fileSystem);
		System.out.println("当前Excel有：" + hssf.getNumberOfSheets() + "个Seelt");
	}
	
	/**
	 * @Description 获取当前Excel中Sheet的内容
	 *
	 * @author 高尚
	 * @version 1.0
	 * @throws IOException 
	 * @date 创建时间：2018年1月26日 下午12:59:32
	 */
	@Test
	public void readSheetValue() throws IOException{
		POIFSFileSystem fileSystem = new POIFSFileSystem(fileIn);
		HSSFWorkbook hssf = new HSSFWorkbook(fileSystem);
		for (int i=0; i<hssf.getNumberOfSheets(); i++) {
			String sheetName = hssf.getSheetName(i);//获取当前Sheet的名称
			System.out.println("工作簿名称："+sheetName);
			HSSFSheet sheet = hssf.getSheetAt(i);//获得当前Sheet对象
			if(null == sheet){
				continue;
			}
			int firstRowNum = sheet.getFirstRowNum();//获取第一个Row
			int lastRowNum = sheet.getLastRowNum();//获取最后一个Row
			for (int j = firstRowNum; j <= lastRowNum; j++) {
				HSSFRow row = sheet.getRow(j);
				if(null == row){
					continue;
				}
				int firstCellNum = row.getFirstCellNum();//获取第一个Cell
				int lastCellNum = row.getLastCellNum();//获取最后一个Cell
				for (int k = firstCellNum; k <= lastCellNum; k++) {
					HSSFCell cell = row.getCell(k);
					if(null == cell){
						continue;
					}
					System.out.print(getCellValue(cell).toString()+"   ");
				}
				System.out.println();
			}
		}
	}
	
	/**
	 * @Description 读取Cell中的内容
	 * @return
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 下午1:32:31
	 */
	private Object getCellValue(HSSFCell cell){
		switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_BOOLEAN:
				return cell.getBooleanCellValue();
			case HSSFCell.CELL_TYPE_NUMERIC:
				return cell.getNumericCellValue();
			default:
				return cell.getStringCellValue();
		}
	}
	
	/**
	 * @Description 抽取Excel中的内容
	 * @throws IOException
	 *
	 * @author 高尚
	 * @version 1.0
	 * @date 创建时间：2018年1月26日 下午1:31:49
	 */
	@Test
	public void readExcelExtractor() throws IOException{
		POIFSFileSystem fileSystem = new POIFSFileSystem(fileIn);
		HSSFWorkbook hssf = new HSSFWorkbook(fileSystem);
		//Excel抽取表面内容
		ExcelExtractor excelExtractor = new ExcelExtractor(hssf);
		excelExtractor.setIncludeSheetNames(false);//不显示Sheet的名称
		System.out.println(excelExtractor.getText());
	}
	
	@After
	public void afterTest() throws IOException{
		if(null != fileIn){
			fileIn.close();
		}
	}

}
