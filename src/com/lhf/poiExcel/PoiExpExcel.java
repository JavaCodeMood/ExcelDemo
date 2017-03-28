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
 * POI����Excel�ļ�
 * @author lhf
 *
 */
public class PoiExpExcel {

	public static void main(String[] args) {
		//�������飬��ű�ͷ��Ϣ
		String[] title = {"id","name","sex","age"};
		//����Excel������
		HSSFWorkbook workbook = new HSSFWorkbook();
		//����һ��������sheet
		HSSFSheet sheet = workbook.createSheet();
		//������һ��  �ļ�ͷ
		HSSFRow row = sheet.createRow(0);
		//���嵥Ԫ��
		HSSFCell cell = null;
		//�ڵ�һ��д��id��name,sex,age
		for(int i=0;i<title.length;i++){
			cell = row.createCell(i);
			cell.setCellValue(title[i]);
		}
		//׷������
		for(int i=1;i<10;i++){
			//�����ڶ���
			HSSFRow nextrow = sheet.createRow(i);
			//���嵥Ԫ��
			HSSFCell cell2 = nextrow.createCell(0);
			//Ϊ��Ԫ��ֵ
			cell2.setCellValue("a"+i);
			cell2 = nextrow.createCell(1);
			cell2.setCellValue("user"+i);
			cell2 = nextrow.createCell(2);
			cell2.setCellValue("Ů");
			cell2 = nextrow.createCell(3);
			cell2.setCellValue("1"+i);
			cell2 = nextrow.createCell(4);
 		}
		//����һ���ļ�������������ɵ�Excel����
		File file = new File("e:/poi_test.xls");
		try {
			//�����ļ�
			file.createNewFile();
			//��Excel���ݴ���
			FileOutputStream stream = FileUtils.openOutputStream(file);
			//д���ļ�
			workbook.write(stream);
			//�ر���
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
       System.out.println("POI�����ļ��ɹ���");
	}

}
