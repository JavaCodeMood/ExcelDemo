package com.lhf.jxlexcel;

import java.io.File;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * JXL创建Excel
 * @author lhf
 *
 */
public class JxlExcel {
	public static void main(String[] args) {
		//用数组存表头
		String[] title={"id","name","sex","age"};
		
		//创建Excel文件
		File file = new File("e:/jxl_test.xls");
		try {
			file.createNewFile();
			//创建工作簿
			WritableWorkbook workbook = Workbook.createWorkbook(file);
			//创建sheet
			WritableSheet sheet = workbook.createSheet("sheet1", 0);
			//往sheet中添加数据
			Label label = null;
			//第一行设置列名
			for(int i=0;i<title.length;i++){
				//Label(i,0,title[i]) 表示第i列第0行，值为title[i]
				label = new Label(i,0,title[i]);
				//添加单元格
				sheet.addCell(label);
			}
			//追加数据
			for(int i=1;i<10;i++){
				//Label(0,i,"a"+1) 表示第0列，第i行，值为“a”+1
				label = new Label(0,i,"a"+i);
				sheet.addCell(label);
				label = new Label(1,i,"user"+i);
				sheet.addCell(label);
				label = new Label(2,i,"男");
				sheet.addCell(label);
				label = new Label(3,i,"20");
				sheet.addCell(label);
			}
			//写入数据
			workbook.write();
			//关闭流
			workbook.close();
			System.out.println("Excel文件创建成功！");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
