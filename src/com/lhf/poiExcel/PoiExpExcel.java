package com.lhf.poiExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * POI生成Excel文件
 * @author lhf
 *
 */
public class PoiExpExcel {

	public static void main(String[] args) {
		//创建数组，存放表头信息
		String[] title = {"id","name","sex","age"};
		//创建Excel工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		//创建一个工作表sheet
		HSSFSheet sheet = workbook.createSheet();
		//创建第一行  文件头
		HSSFRow row = sheet.createRow(0);
		//定义单元格
		HSSFCell cell = null;
		//在第一行写入id，name,sex,age
		for(int i=0;i<title.length;i++){
			cell = row.createCell(i);
			cell.setCellValue(title[i]);
		}
		//追加数据
		for(int i=1;i<10;i++){
			//创建第二行
			HSSFRow nextrow = sheet.createRow(i);
			//定义单元格
			HSSFCell cell2 = nextrow.createCell(0);
			//为单元格赋值
			cell2.setCellValue("a"+i);
			cell2 = nextrow.createCell(1);
			cell2.setCellValue("user"+i);
			cell2 = nextrow.createCell(2);
			cell2.setCellValue("女");
			cell2 = nextrow.createCell(3);
			cell2.setCellValue("1"+i);
			cell2 = nextrow.createCell(4);
 		}
		//创建一个文件，用来存放生成的Excel数据
		File file = new File("e:/poi_test.xls");
		try {
			//创建文件
			file.createNewFile();
			//将Excel内容存盘
			FileOutputStream stream = FileUtils.openOutputStream(file);
			//写入文件
			workbook.write(stream);
			//关闭流
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
       System.out.println("POI创建文件成功！");
	}

}
