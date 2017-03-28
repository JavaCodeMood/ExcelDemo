package com.lhf.excel_template;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;

/**
 * ����ģ���ļ�
 * @author lhf
 *
 */
public class CreateTemplate {
	public static void main(String[] args) {
		//��ȡ����xml�ļ�·��
		String path = System.getProperty("user.dir")+"/bin/student.xml";
	    System.out.println(path);
	    File file = new File(path);
	    //��������
	    SAXBuilder builder = new SAXBuilder();
	    try {
	    	//����xml�ļ�
			Document parse = builder.build(file);
			//����һ��Excel
			//����������
			HSSFWorkbook wb = new HSSFWorkbook();
			//����sheet
			HSSFSheet sheet = wb.createSheet("Sheet0");
			
			//��ȡxml�ļ����ڵ�  <excel></excel>
			Element root = parse.getRootElement();
			//��ȡģ�������
			String templateName = root.getAttribute("name").getValue();
			int rownum = 0;  //�к�
			int colnum = 0;   //�к�
			//�����п�
			Element colgroup = root.getChild("colgroup");
			setColumnWidth(sheet,colgroup);
			
			/*���ñ��� �ϲ���Ԫ��
			 *<title>
             *<tr height="16px">
             *<td rowspan="1" colspan="6" value="ѧ����Ϣ����" />
             *</tr>
             *</title>
			 */
			Element title = root.getChild("title");   //��ȡtitle��ǩ<title></title>
			//��ȡtr��ǩ
			List<Element> trs = title.getChildren("tr");
			//ѭ��tr
			for(int i=0;i<trs.size();i++){
				Element tr = trs.get(i);   //��ȡtr
				List<Element> tds = tr.getChildren("td");  //��ȡtr�е�td
				//����һ��
				HSSFRow row = sheet.createRow(rownum);
				//Ϊ��Ԫ��������ʽ--����
				HSSFCellStyle cellStyle = wb.createCellStyle();
				cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				//�����ж��td
				for(colnum=0;colnum<tds.size();colnum++){
					Element td = tds.get(colnum);
					//������Ԫ��
					HSSFCell cell = row.createCell(colnum);
					//��ȡ��������
					Attribute rowSpan = td.getAttribute("rowspan");
					Attribute colSpan = td.getAttribute("colspan");
					Attribute value = td.getAttribute("value");
				    //���value������null
					if(value != null){
						String val = value.getValue();
						//��ֵ���뵥Ԫ��
						cell.setCellValue(val);
						//��ʼ��
						int rspan = rowSpan.getIntValue()-1;
						//�ϲ���
						int cspan = colSpan.getIntValue()-1;
						//��������
						HSSFFont font =wb.createFont();
						font.setFontName("���Ρ���GB2312");
						//����Ӵ�
						font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
						//��ɫ����
						font.setColor(HSSFFont.COLOR_RED);
						//����߶�
						//font.setFontHeight((short)12);
						font.setFontHeightInPoints((short)12);
						//���������ʽ��
						cellStyle.setFont(font);
						//Ϊ��Ԫ��������ʽ
						cell.setCellStyle(cellStyle);
						
						//�ϲ���Ԫ��
						sheet.addMergedRegion(new CellRangeAddress(rspan,rspan,0,cspan));
					}
				}
				rownum++;
			}
			
			//���ñ�ͷ��Ϣ
			/* <thead>
               <tr height="16px">
        	   <th value="���" />
               <th value="����" />
               <th value="����" />
               <th value="�Ա�" />
               <th value="��������" />
               <th value=" ����" />            
               </tr>
               </thead>
            */
			//���ñ�ͷ
			Element thead = root.getChild("thead");
			trs = thead.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				//����һ��
				HSSFRow row = sheet.createRow(rownum);
				//��ȡth�Ľڵ���Ϣ
				List<Element> ths = tr.getChildren("th");
				for(colnum = 0;colnum < ths.size();colnum++){
					//ȡ��thԪ��
					Element th = ths.get(colnum);
					//���th������
					Attribute valueAttr = th.getAttribute("value");
					//������Ԫ��
					HSSFCell cell = row.createCell(colnum);
					if(valueAttr != null){
						String value =valueAttr.getValue();
						//��ֵ
						cell.setCellValue(value);
					}
				}
				rownum++;
			}
			
			//��������������ʽ
			Element tbody = root.getChild("tbody");  //��ȡtbody�ڵ�
			Element tr = tbody.getChild("tr");  //��ȡtr�ڵ�
			//��ȡrepeat����
			int repeat = tr.getAttribute("repeat").getIntValue();
			//��ȡtd
			List<Element> tds = tr.getChildren("td");
			for (int i = 0; i < repeat; i++) {
				//�����м�¼
				HSSFRow row = sheet.createRow(rownum);
				//���õ�Ԫ��
				for(colnum =0 ;colnum < tds.size();colnum++){
					//��ȡtd�ڵ�
					Element td = tds.get(colnum);
					//������Ԫ��
					HSSFCell cell = row.createCell(colnum);
					//���õ�Ԫ����ʽ
					setType(wb,cell,td);
				}
				rownum++;
			}
			//��Excel�ļ����浽����
			File tempFile = new File("e:/"+templateName + ".xls");
			tempFile.delete();  //���ھ�ɾ��
			//�����ھʹ���
			tempFile.createNewFile();
			FileOutputStream stream = FileUtils.openOutputStream(tempFile);
			//д��
			wb.write(stream);
			//�ر�
			stream.close();
			
			
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

	/**
	 * ���õ�Ԫ����ʽ
	 * @param wb
	 * @param cell
	 * @param td
	 */
	private static void setType(HSSFWorkbook wb, HSSFCell cell, Element td) {
		Attribute typeAttr = td.getAttribute("type");
		String type = typeAttr.getValue();
		HSSFDataFormat format = wb.createDataFormat();
		HSSFCellStyle cellStyle = wb.createCellStyle();
		//�жϽڵ�����
		if("NUMERIC".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			Attribute formatAttr = td.getAttribute("format");
			String formatValue = formatAttr.getValue();
			formatValue = StringUtils.isNotBlank(formatValue)? formatValue : "#,##0.00";
			cellStyle.setDataFormat(format.getFormat(formatValue));
		}else if("STRING".equalsIgnoreCase(type)){
			cell.setCellValue("");
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cellStyle.setDataFormat(format.getFormat("@"));
		}else if("DATE".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cellStyle.setDataFormat(format.getFormat("yyyy-m-d"));
		}else if("ENUM".equalsIgnoreCase(type)){
			CellRangeAddressList regions = 
				new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(), 
						cell.getColumnIndex(), cell.getColumnIndex());
			Attribute enumAttr = td.getAttribute("format");
			String enumValue = enumAttr.getValue();
			//���������б�����
			DVConstraint constraint = 
				DVConstraint.createExplicitListConstraint(enumValue.split(","));
			//������Ч�Զ���
			HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
			wb.getSheetAt(0).addValidationData(dataValidation);
		}
		cell.setCellStyle(cellStyle);
	}
	/**
	 * �����п�
	 * @param sheet
	 * @param colgroup
	 */
	private static void setColumnWidth(HSSFSheet sheet, Element colgroup) {
		//���ÿһ������
		List<Element> cols = colgroup.getChildren("col");
		for(int i=0;i<cols.size();i++){
			//��ȡÿһ��col������
			Element col = cols.get(i);
			Attribute width = col.getAttribute("width");
			//�õ���λem
			String unit = width.getValue().replaceAll("[0-9,\\.]", "");
			//�õ�ֵ
			String value = width.getValue().replaceAll(unit, "");
			int v = 0;
			//�����λΪ�ջ���Ϊpx
			if(StringUtils.isBlank(unit) || "px".endsWith(unit)){
				//��poi�Ŀ��ת��ΪExcel�Ŀ��
				v = Math.round(Float.parseFloat(value)*37F);
			}else if("em".endsWith(unit)){  //�����λΪem
				v = Math.round(Float.parseFloat(value)*267.5F);
			}
			//���ÿ��
			sheet.setColumnWidth(i, v);
			
			
		}
		
	}

}
