package com.hpugs.poi.servlet;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.hpugs.poi.entity.User;

/**
 * 导出Excel
 */
@WebServlet("/ExportExcel")
public class ExportExcel extends HttpServlet {
	
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.setHeader("Content-disposition", "attachment; filename="+ new String("单元格导出测试".getBytes(),"ISO8859-1") +".xls");// 设定输出文件头   
        response.setContentType("application/msexcel");// 定义输出类型 
		OutputStream output = response.getOutputStream();//得到输出流
		//创建工作簿
		Workbook wb = new HSSFWorkbook();
		
		//Sheet页标题
		String[] heads = {"Id", "真实姓名", "昵称", "年龄", "存款", "创建时间", "备注"};
		wb = createSheet(wb, heads, 100);
		wb = createSheet(wb, heads, 340);
		wb = createSheet(wb, heads, 80);
		wb = createSheet(wb, heads, 2000);
		
		wb.write(output);
		if(null != output){
			output.flush();
			output.close();
		}
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}
	
	/**
	 * 生成Sheet页
	 * @param wb
	 * @param heads
	 * @param userCount
	 * @return
	 */
	private Workbook createSheet(Workbook wb, String[] heads, int userCount){
		//第一个Sheet页
		Sheet sheet = wb.createSheet(userCount+"条记录");
		int rowIndex = 0;
		//创建标题
		Row row = sheet.createRow(rowIndex++);
		for (int i=0; i<heads.length; i++) {
			Cell cell = row.createCell(i);
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
			cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
			cell.setCellStyle(cellStyle);
			cell.setCellValue(heads[i]);
		}
		//创建内容
		List<User> users = User.createUserList(userCount);
		//设置单元格样式
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(CellStyle.ALIGN_RIGHT);//水平居右
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
		cellStyle.setWrapText(true);//自动换行
		//设置金额格式化
		DataFormat format = wb.createDataFormat();
		CellStyle numberCellStyle = wb.createCellStyle();
		numberCellStyle.setDataFormat(format.getFormat("#,##0.00"));
		//设置时间格式化
		CellStyle dateCellStyle = wb.createCellStyle();
		CreationHelper creationHelper = wb.getCreationHelper();
		dateCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
		
		//遍历用户数据
		for (User user : users) {
			row = sheet.createRow(rowIndex++);
			Cell cell = row.createCell(0);
			cell.setCellValue(user.getId());
			cell.setCellStyle(cellStyle);
			
			cell = row.createCell(1);
			cell.setCellValue(user.getRealName());
			cell.setCellStyle(cellStyle);
			
			cell = row.createCell(2);
			cell.setCellValue(user.getNickName());
			cell.setCellStyle(cellStyle);
			
			cell = row.createCell(3);
			cell.setCellValue(user.getAge());
			cell.setCellStyle(cellStyle);
			
			cell = row.createCell(4);
			cell.setCellValue(user.getMonery());
			cell.setCellStyle(numberCellStyle);
			
			cell = row.createCell(5);
			cell.setCellValue(user.getGmtCreat());
			cell.setCellStyle(dateCellStyle);
			
			cell = row.createCell(6);
			cell.setCellValue(user.getRemark());
			cell.setCellStyle(cellStyle);
		}
		return wb;
	}

}
