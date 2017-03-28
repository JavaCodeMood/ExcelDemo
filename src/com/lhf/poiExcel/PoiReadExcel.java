package com.lhf.poiExcel;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * POI����Excel�ļ�
 * @author lhf
 *
 */
public class PoiReadExcel {

	public static void main(String[] args) {
		//��Ҫ������Excel�ļ�
		File file = new File("e:/poi_test.xls");
		try {
			//��������������ȡ�ļ�����
			HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
			//��ȡExcel�ļ��ĵ�һ������ҳ
			//��ʽһ
			//HSSFSheet sheet = workbook.getSheet("sheet0");
			//��ʽ��
			HSSFSheet sheet = workbook.getSheetAt(0);
			int firstRowNum = 0;  //��һ��
			//��ȡsheet�����һ���к�
			int lastRowNum = sheet.getLastRowNum();  //��ȡ���һ��
			for(int i=0;i<=lastRowNum;i++){
				HSSFRow row = sheet.getRow(i);
				//��ȡ��ǰ�����Ԫ���к�
				int lastCellNum = row.getLastCellNum();
				//ѭ����һ������ȡÿһ����Ԫ���е�����
				for(int j=0;j<lastCellNum;j++){
					//��ȡ
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
