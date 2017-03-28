package com.lhf.poiExcel;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * POI解析Excel文件
 * @author lhf
 *
 */
public class PoiReadExcel {

	public static void main(String[] args) {
		//需要解析的Excel文件
		File file = new File("e:/poi_test.xls");
		try {
			//创建工作簿，读取文件内容
			HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
			//读取Excel文件的第一个工作页
			//方式一
			//HSSFSheet sheet = workbook.getSheet("sheet0");
			//方式二
			HSSFSheet sheet = workbook.getSheetAt(0);
			int firstRowNum = 0;  //第一行
			//获取sheet中最后一行行号
			int lastRowNum = sheet.getLastRowNum();  //获取最后一行
			for(int i=0;i<=lastRowNum;i++){
				HSSFRow row = sheet.getRow(i);
				//获取当前行最后单元格列号
				int lastCellNum = row.getLastCellNum();
				//循环这一行来读取每一个单元格中的内容
				for(int j=0;j<lastCellNum;j++){
					//读取
					HSSFCell cell = row.getCell(j);
					String value = cell.getStringCellValue();
					System.out.print(value+" ");
					
				}
				System.out.println();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		

	}

}
