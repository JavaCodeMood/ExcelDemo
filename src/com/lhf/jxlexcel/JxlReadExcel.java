package com.lhf.jxlexcel;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableWorkbook;

/**
 * JXL解析Excel
 * @author lhf
 *
 */
public class JxlReadExcel {

	public static void main(String[] args) {
		//File file = new File("e:/jxl_test.xls");
		try {
			//创建工作簿
			Workbook workbook = Workbook.getWorkbook(new File("e:/jxl_test.xls"));
			//获取第一个工作表sheet页
			Sheet sheet = workbook.getSheet(0);
			//循环获取
			//1.循环行
			for(int i=0;i<sheet.getRows();i++){
				//2.循环列
				for(int j=0;j<sheet.getColumns();j++){
					//获取单元格内容
					Cell cell = sheet.getCell(j,i);
					System.out.print(cell.getContents()+" ");
				}
				System.out.println();
			}
			//关闭流
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
