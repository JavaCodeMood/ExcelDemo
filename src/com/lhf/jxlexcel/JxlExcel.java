package com.lhf.jxlexcel;

import java.io.File;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * JXL����Excel
 * @author lhf
 *
 */
public class JxlExcel {
	public static void main(String[] args) {
		//��������ͷ
		String[] title={"id","name","sex","age"};
		
		//����Excel�ļ�
		File file = new File("e:/jxl_test.xls");
		try {
			file.createNewFile();
			//����������
			WritableWorkbook workbook = Workbook.createWorkbook(file);
			//����sheet
			WritableSheet sheet = workbook.createSheet("sheet1", 0);
			//��sheet���������
			Label label = null;
			//��һ����������
			for(int i=0;i<title.length;i++){
				//Label(i,0,title[i]) ��ʾ��i�е�0�У�ֵΪtitle[i]
				label = new Label(i,0,title[i]);
				//��ӵ�Ԫ��
				sheet.addCell(label);
			}
			//׷������
			for(int i=1;i<10;i++){
				//Label(0,i,"a"+1) ��ʾ��0�У���i�У�ֵΪ��a��+1
				label = new Label(0,i,"a"+i);
				sheet.addCell(label);
				label = new Label(1,i,"user"+i);
				sheet.addCell(label);
				label = new Label(2,i,"��");
				sheet.addCell(label);
				label = new Label(3,i,"20");
				sheet.addCell(label);
			}
			//д������
			workbook.write();
			//�ر���
			workbook.close();
			System.out.println("Excel�ļ������ɹ���");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
