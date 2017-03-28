package com.lhf.jxlexcel;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableWorkbook;

/**
 * JXL����Excel
 * @author lhf
 *
 */
public class JxlReadExcel {

	public static void main(String[] args) {
		//File file = new File("e:/jxl_test.xls");
		try {
			//����������
			Workbook workbook = Workbook.getWorkbook(new File("e:/jxl_test.xls"));
			//��ȡ��һ��������sheetҳ
			Sheet sheet = workbook.getSheet(0);
			//ѭ����ȡ
			//1.ѭ����
			for(int i=0;i<sheet.getRows();i++){
				//2.ѭ����
				for(int j=0;j<sheet.getColumns();j++){
					//��ȡ��Ԫ������
					Cell cell = sheet.getCell(j,i);
					System.out.print(cell.getContents()+" ");
				}
				System.out.println();
			}
			//�ر���
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
